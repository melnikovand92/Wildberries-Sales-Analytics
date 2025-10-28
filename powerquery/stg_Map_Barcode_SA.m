let
  // === Бери готовый факт-заказы как источник ===
  Source = f_Orders,

  // === Оставляем только нужные поля ===
  Keep = Table.SelectColumns(Source, {"barcode", "supplierArticle"}),

  // === Типы ===
  Typed = Table.TransformColumnTypes(Keep, {{"barcode", type text}, {"supplierArticle", type text}}),

  // === Чистка barcode ===
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

  // === Нормализация supplierArticle → supplierArticle_norm ===
  AddSupplierNorm = Table.AddColumn(
                      CleanBarcode, "supplierArticle_norm",
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

  // === Убираем пустые ===
  NoNulls = Table.SelectRows(AddSupplierNorm, each [barcode] <> null and [supplierArticle_norm] <> null),

  // === Делаем маппинг без дублей по barcode ===
  DistinctMap = Table.Distinct(NoNulls, {"barcode"}),

  // === Возвращаем только нужные столбцы карты ===
  Result = Table.SelectColumns(DistinctMap, {"barcode", "supplierArticle_norm"})
in
  Result
