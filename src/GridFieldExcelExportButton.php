<?php

namespace Level51\ExcelExport;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SilverStripe\Control\Controller;
use SilverStripe\Control\HTTPRequest;
use SilverStripe\Core\Config\Config;
use SilverStripe\Forms\GridField\GridField;
use SilverStripe\Forms\GridField\GridField_ActionProvider;
use SilverStripe\Forms\GridField\GridField_FormAction;
use SilverStripe\Forms\GridField\GridField_HTMLProvider;
use SilverStripe\Forms\GridField\GridField_URLHandler;
use SilverStripe\Forms\GridField\GridFieldFilterHeader;
use SilverStripe\Forms\GridField\GridFieldSortableHeader;
use SilverStripe\Forms\GridField\GridFieldPaginator;

class GridFieldExcelExportButton implements GridField_HTMLProvider, GridField_ActionProvider, GridField_URLHandler {

    /**
     * @var array Map of a property name on the exported objects, with values being the column title in the CSV file.
     * Note that titles are only used when {@link $csvHasHeader} is set to TRUE.
     */
    protected $exportColumns;

    protected $targetFragment;

    /**
     * @var callable
     */
    protected $afterExportCallback;

    public function __construct($targetFragment = 'before', $exportColumns = null) {
        $this->targetFragment = $targetFragment;
        $this->exportColumns = $exportColumns;
    }

    public function getHTMLFragments($gridField) {
        $button = new GridField_FormAction(
            $gridField,
            'exportexcel',
            _t(__CLASS__ . '.EXPORT_CTA', 'Export as Excel file'),
            'exportexcel',
            null
        );
        $button->addExtraClass('btn btn-secondary no-ajax font-icon-down-circled action_export');
        $button->setForm($gridField->getForm());

        return [
            $this->targetFragment => $button->Field(),
        ];
    }


    public function getActions($gridField) {
        return ['exportexcel'];
    }


    public function handleAction(GridField $gridField, $actionName, $arguments, $data) {
        if ($actionName == 'exportexcel') {
            return $this->handleExcelExport($gridField);
        }
    }

    public function getURLHandlers($gridField) {
        return [
            'exportexcel' => 'handleExcelExport',
        ];
    }

    protected function getExportColumnsForGridField(GridField $gridField) {
        if ($this->exportColumns) {
            $exportColumns = $this->exportColumns;
        } else if ($dataCols = $gridField->getConfig()->getComponentByType('GridFieldDataColumns')) {
            $exportColumns = $dataCols->getDisplayFields($gridField);
        } else {
            $exportColumns = singleton($gridField->getModelClass())->summaryFields();
        }

        return $exportColumns;
    }

    public function handleExcelExport($gridField, $request = null) {

        // Setup filename and path (to tmp dir)
        $now = Date("d-m-Y-H-i");
        $fileName = "excel-export-$now.xlsx";
        $path = Controller::join_links(sys_get_temp_dir(), $fileName);

        // Setup spreadsheet
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $columns = $this->getExportColumnsForGridField($gridField);

        // Get headers (@see GridFieldExportButton)
        $headers = [];
        foreach ($columns as $columnSource => $columnHeader) {
            if (is_array($columnHeader) && array_key_exists('title', $columnHeader)) {
                $headers[] = $columnHeader['title'];
            } else {
                $headers[] = (!is_string($columnHeader) && is_callable($columnHeader)) ? $columnSource : $columnHeader;
            }
        }

        // Write headers
        $sheet->fromArray($headers, null, 'A1');

        //Remove GridFieldPaginator as we're going to export the entire list.
        $gridField->getConfig()->removeComponentsByType(GridFieldPaginator::class);

        // Collect items / content rows (@see GridFieldExportButton)
        $items = $gridField->getManipulatedList();

        // @todo should GridFieldComponents change behaviour based on whether others are available in the config?
        foreach ($gridField->getConfig()->getComponents() as $component) {
            if ($component instanceof GridFieldFilterHeader || $component instanceof GridFieldSortableHeader) {
                $items = $component->getManipulatedData($gridField, $items);
            }
        }

        $content = [];
        foreach ($items->limit(null) as $item) {
            if (!$item->hasMethod('canView') || $item->canView()) {
                $columnData = array();

                foreach ($columns as $columnSource => $columnHeader) {
                    if (!is_string($columnHeader) && is_callable($columnHeader)) {
                        if ($item->hasMethod($columnSource)) {
                            $relObj = $item->{$columnSource}();
                        } else {
                            $relObj = $item->relObject($columnSource);
                        }

                        $value = $columnHeader($relObj);
                    } else {
                        $value = $gridField->getDataFieldValue($item, $columnSource);

                        if ($value === null) {
                            $value = $gridField->getDataFieldValue($item, $columnHeader);
                        }
                    }

                    $value = str_replace(array("\r", "\n"), "\n", $value);

                    // [SS-2017-007] Sanitise XLS executable column values with a leading tab
                    if (!Config::inst()->get(get_class($this), 'xls_export_disabled')
                        && preg_match('/^[-@=+].*/', $value)
                    ) {
                        $value = "\t" . $value;
                    }
                    $columnData[] = $value;
                }

                $content[] = $columnData;
            }

            if ($item->hasMethod('destroy')) {
                $item->destroy();
            }
        }

        // Write content rows
        $sheet->fromArray($content, null, 'A2');

        // Auto size column width
        $headerRow = $sheet->getRowIterator(1,1)->current();

        foreach ($headerRow->getCellIterator() as $cell) {
            $sheet->getColumnDimension($cell->getColumn())
                ->setAutoSize(true);
        }

        // Write xlsx file
        $writer = new Xlsx($spreadsheet);

        if ($callback = $this->getAfterExportCallback()) {
            $callback($writer);
        }

        $writer->save($path);

        // Read and return file content (triggers download) (MIME Type: https://blogs.msdn.microsoft.com/vsofficedeveloper/2008/05/08/office-2007-file-format-mime-types-for-http-content-streaming-2/)
        return HTTPRequest::send_file(file_get_contents($path), $fileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }

    /**
     *
     * @return callable
     */
    public function getAfterExportCallback()
    {
        return $this->afterExportCallback;
    }

    /**
     *
     * @param callable $afterExportCallback
     * @return ExcelGridFieldExportButton
     */
    public function setAfterExportCallback(callable $afterExportCallback)
    {
        $this->afterExportCallback = $afterExportCallback;
        return $this;
    }
}
