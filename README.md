# SilverStripe GridField button for excel exports

Similar to the default GridFieldExportButton but exports Excel (xlsx) files instead of CSV. 

## Requirements
- SilverStripe ^4.0 (see 0.1.0 tag for SilverStripe 3 support)
- php >= 7.1
- phpoffice/phpspreadsheet ^1.11

## Howto

### Styling the exported xlsx file

You can hook in and style the ready-to-export Excel file with a callback function.

```php
$excelExportButton = new GridFieldExcelExportButton('buttons-before-left', $exportFieldMapping);
        $excelExportButton->setAfterExportCallback([ExcelStylingHelper::class, 'styleExcelExport']);
```

The callback passes a subclass of PhpOffice\PhpSpreadsheet\Writer\BaseWriter (e.g. Xlsx writer object), there you can get the spreadsheet and the worksheet and manipulate it.  [See PHPSpreadsheet docs fo more details](https://phpspreadsheet.readthedocs.io/en/latest/).


### Setting the default font

```php

class ExcelStylingHelper
{
    public static function styleExcelExport(BaseWriter $writer): void
    {
        $sheet = $writer->getSpreadsheet();
        $sheet->getDefaultStyle()->getFont()->setName('Comic Sans MS');
        $sheet->getDefaultStyle()->getFont()->setSize(12);
    }
}
```

### Draw borders around the exported data
and a thicker line between header and data
```php
        /** @var Worksheet $worksheet */
        $worksheet = $sheet->getActiveSheet();
        $borders = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => 'FF000000']
                ]
            ]
        ];
        $bottomRight = $worksheet->getHighestColumn() . $worksheet->getHighestRow();
        $worksheet->getStyle('A1:' . $bottomRight)->applyFromArray($borders);
        $worksheet->getStyle('A1:' . $worksheet->getHighestColumn() . '1')->applyFromArray(
            [
                'borders' => [
                    'bottom' => [
                        'borderStyle' => Border::BORDER_THICK,
                        'color' => ['argb' => 'FF000000']
                    ]
                ]
            ]
        );
```

### Style a specific column
Sometimes, e.g. telephone numbers are displayed as scientific number. In this case you need to format the content as text.

In this example we check the header row for a specific column named "Tel" and format each cell in this column explicitely as text.

```php
        $headerRow = $worksheet->getRowIterator(1, 1)->current();

        foreach ($headerRow->getCellIterator() as $headerCell) {
            //check if header cell is "Tel"
            if ($headerCell->getValue() === 'Tel') {
                //format column as text
                $column = $headerCell->getColumn();
                $worksheet->getStyle($column . ':' . $column)
                    ->getNumberFormat()
                    ->setFormatCode(NumberFormat::FORMAT_TEXT);
                $column = $worksheet->getColumnIterator($column, $column)->current();
                foreach ($column->getCellIterator() as $valueCell) {
                    $valueCell->setValueExplicit(
                        $valueCell->getValue(),
                        DataType::TYPE_STRING
                    );
                }
            }

        }
```

## Maintainer
- Daniel Kliemsch <dk@lvl51.de>
- Julian Scheuchenzuber <js@lvl51.de>