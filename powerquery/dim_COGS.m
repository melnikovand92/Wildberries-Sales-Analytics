let
  // === Источник ===
  Url = "https://docs.google.com/spreadsheets/d/10sH4PxnIeT2AO8RK-b6SpHB8TuBfbXwC/export?format=xlsx",
  Bin = Web.Contents(Url),
  Xls = Excel.Workbook(Bin, null, true),
  Pick = Table.SelectRows(Xls, each [Kind] = "Sheet" and Text.Lower(Text.Trim([Item])) = "себестоимость"){0}[Data],
  Promoted = Table.PromoteHeaders(Pick, [PromoteAllScalars = true]),

  // === Типы ===
  Typed = Table.TransformColumnTypes(
            Promoted,
            {
              {"Штрих-код", type text},
              {"Себестоимость", type number}
            }
          ),

  // === Чистка Штрих-код ===
  CleanBarcode = Table.TransformColumns(
                   Typed,
                   {
                     {"Штрих-код", each
                        let
                          t0 = if _ = null then null else Text.Clean(Text.Trim(_)),
                          t1 = if t0 = null then null else Text.Replace(t0, Character.FromNumber(160), " "),
                          t2 = if t1 = null then null else Text.Replace(t1, Character.FromNumber(8239), " ")
                        in t2,
                      type text}
                   }
                 ),

  // === Merge с dim_Product по barcode, чтобы получить supplierArticle_norm ===
  Merged = Table.NestedJoin(
             CleanBarcode,
             {"Штрих-код"},
             dim_Product,
             {"barcode"},
             "prod",
             JoinKind.LeftOuter
           ),
  Expanded = Table.ExpandTableColumn(Merged, "prod", {"supplierArticle_norm"}, {"supplierArticle_norm"}),

  // === Диагностика покрытия ===
  AddDiag = Table.AddColumn(
              Expanded, "для диагностики",
              each if [supplierArticle_norm] = null then "NO MATCH" else "MATCH",
              type text
            ),

  // === Опционально убрать явные дубли по Штрих-код ===
  DistinctByBarcode = Table.Distinct(AddDiag, {"Штрих-код"})
in
  DistinctByBarcode
