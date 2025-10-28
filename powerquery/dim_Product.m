let
  // === Источник (Google Sheets → Excel.Workbook(Web.Contents)) ===
  Url = "https://docs.google.com/spreadsheets/d/10sH4PxnIeT2AO8RK-b6SpHB8TuBfbXwC/export?format=xlsx",
  Bin = Web.Contents(Url),
  Xls = Excel.Workbook(Bin, null, true),
  Pick = Table.SelectRows(Xls, each [Kind] = "Sheet" and Text.Lower(Text.Trim([Item])) = "справочник номенклатуры"){0}[Data],
  Promoted = Table.PromoteHeaders(Pick, [PromoteAllScalars = true]),

  // === Типы (минимум нужного) ===
  Typed = Table.TransformColumnTypes(
            Promoted,
            {
              {"supplierArticle", type text},
              {"barcode", type text},
              {"subject", type text},
              {"category", type text},
              {"brand", type text},
              {"Рекомендуемая цена", type number}
            }
          ),

  // === Чистка barcode: Trim/Clean + замена NBSP (U+00A0, U+202F) на обычный пробел, верхний регистр опционально ===
  CleanBarcode = Table.TransformColumns(
                   Typed,
                   {
                     {"barcode", each
                        let
                          t0 = if _ = null then null else Text.Clean(Text.Trim(_)),
                          t1 = if t0 = null then null else Text.Replace(t0, Character.FromNumber(160), " "),
                          t2 = if t1 = null then null else Text.Replace(t1, Character.FromNumber(8239), " ")
                        in t2,
                      type text}
                   }
                 ),

  // === Нормализация supplierArticle → supplierArticle_norm (NBSP→пробел, Trim/Clean, collapse spaces, Upper) ===
  AddSupplierNorm = Table.AddColumn(
                      CleanBarcode,
                      "supplierArticle_norm",
                      each
                        let
                          raw = try Text.From([supplierArticle]) otherwise null,
                          t0  = if raw = null then null else Text.Clean(Text.Trim(raw)),
                          t1  = if t0  = null then null else Text.Replace(t0, Character.FromNumber(160), " "),
                          t2  = if t1  = null then null else Text.Replace(t1, Character.FromNumber(8239), " "),
                          lst = if t2  = null then {}   else List.Select(Text.Split(t2, " "), each _ <> ""),
                          one = if List.Count(lst) = 0 then null else Text.Combine(lst, " "),
                          res = if one = null then null else Text.Upper(one, "ru-RU")
                        in
                          res,
                      type text
                    ),

  // === Удаляем дубликаты по barcode (патч/исходник перекрывается по первому вхождению) ===
  DistinctByBarcode = Table.Distinct(AddSupplierNorm, {"barcode"})
in
  DistinctByBarcode
