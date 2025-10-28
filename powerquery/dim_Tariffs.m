let
  // === Источник ===
  Url = "https://docs.google.com/spreadsheets/d/10sH4PxnIeT2AO8RK-b6SpHB8TuBfbXwC/export?format=xlsx",
  Bin = Web.Contents(Url),
  Xls = Excel.Workbook(Bin, null, true),
  Pick = Table.SelectRows(Xls, each [Kind] = "Sheet" and Text.Lower(Text.Trim([Item])) = "тарифы и доставка"){0}[Data],
  Promoted = Table.PromoteHeaders(Pick, [PromoteAllScalars = true]),

  // === Типы (подстрой под фактические названия колонок в листе) ===
  Typed = Table.TransformColumnTypes(
            Promoted,
            {
              {"subject", type text},
              {"Процент комиссии", type number},
              {"Процент комиссии по FBS", type number},
              {"Cтоимость логистики", type number},
              {"Базовая стоимость хранения", type number},
              {"Платная приёмка", type number},
              {"₽/ед.товара", type number},
              {"Расчёт по фактическим габаритам товара", type text}
            }
          )
in
  Typed
