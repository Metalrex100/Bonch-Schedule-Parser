<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

class ScheduleExcelExporter
{
    const WEEKDAYS = [
        1 => 'Понедельник',
        2 => 'Вторник',
        3 => 'Среда',
        4 => 'Четверг',
        5 => 'Пятница',
        6 => 'Суббота',
        7 => 'Воскресенье',
    ];

    const COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F'];

    const MAX_PAIRS_COUNT = 5;
    const CELLS_FOR_SINGLE_PAIR = 5;
    const CELLS_FOR_SINGLE_DAY = 26;

    const CELL_COLOR_WEEKDAY = 'ff6666';
    const CELL_COLOR_PAIR_DAY = 'FFDD99';
    const CELL_FONT_SIZE = 14;

    const CELL_TYPE_WEEKDAY = 'weekday';
    const CELL_TYPE_PAIR_NUMBER = 'pair_number';

    /**
     * @var Spreadsheet
     */
    private $spreadsheet;

    /**
     * @var int
     */
    protected $finalRow;

    /**
     * @var string
     */
    protected $firstCell = 'A1';

    /**
     * @var string
     */
    protected $lastCell;

    /**
     * @var int
     */
    protected $extraCells = 0;

    /**
     * ScheduleExcelExporter constructor.
     */
    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
    }

    /**
     * @param array $schedule
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function export(array $schedule)
    {
        $this->fillSheet($schedule);
        $this->setupDocument();
        $this->saveDocument();
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setupDocument()
    {
        $finalCell = sprintf('%s%d', array_last(self::COLUMNS), $this->finalRow);
        $cellsRange = sprintf('%s:%s', $this->firstCell, $finalCell);

        $sheet = $this->spreadsheet->getActiveSheet();
        $sheet
            ->getStyle($cellsRange)
            ->getBorders()
            ->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setRGB('000000')
        ;

        $sheet
            ->getStyle($cellsRange)
            ->getFont()
            ->setSize(self::CELL_FONT_SIZE)
            ->setName(\PhpOffice\PhpSpreadsheet\Shared\Font::TIMES_NEW_ROMAN)
        ;

        $sheet
            ->getStyle($cellsRange)
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ;

        foreach (self::COLUMNS as $column) {
            $sheet->getColumnDimension($column)->setAutoSize(true);
        }
    }

    /**
     * @param array $schedule
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function fillSheet(array $schedule)
    {
        foreach ($schedule as $weekNumber => $weekdays) {
            $this->fillWeekdays($weekNumber, $weekdays);
        }
    }

    /**
     * @param int   $weekNumber
     * @param array $weekdays
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function fillWeekdays(int $weekNumber, array $weekdays)
    {
        $initialRow = $weekNumber * self::CELLS_FOR_SINGLE_DAY + 1;

        foreach ($weekdays as $weekday => $dates) {
            $this->fillDates($initialRow, $weekday, $dates);
        }
    }

    /**
     * @param int   $initialRow
     * @param int   $weekday
     * @param array $dates
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function fillDates(int $initialRow, int $weekday, array $dates)
    {
        $column = $weekday;

        foreach ($dates as $date => $pairs) {
            $this->fillCell(
                $column,
                $initialRow,
                sprintf('%s, %s', self::WEEKDAYS[$weekday], $date),
                self::CELL_TYPE_WEEKDAY
            );

            $this->fillPairs($column, ++$initialRow, $pairs);
        }
    }

    /**
     * @param int   $column
     * @param int   $row
     * @param array $pairs
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function fillPairs(int $column, int $row, array $pairs)
    {
        foreach ($pairs as $pair) {
            $pairRow = ($pair['number'] - 1) * self::CELLS_FOR_SINGLE_PAIR + $row;

            $this->fillCell(
                $column,
                $pairRow,
                sprintf('%d %s', $pair['number'], $pair['time']),
                self::CELL_TYPE_PAIR_NUMBER
            );

            $this->fillCell($column, ++$pairRow, $pair['type']);
            $this->fillCell($column, ++$pairRow, $pair['subject']);
            $this->fillCell($column, ++$pairRow, $pair['teacher']);
            $this->fillCell($column, ++$pairRow, $pair['address']);
        }
    }

    /**
     * @param int    $column
     * @param int    $row
     * @param mixed  $value
     *
     * @param string $type
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function fillCell(int $column, int $row, $value, string $type = null)
    {
        $cell = $this->spreadsheet->getActiveSheet()->getCellByColumnAndRow($column, $row)->setValue($value);

        $this->setCellStyle($cell, $type);
    }

    /**
     * @param Cell   $cell
     * @param string $type
     */
    protected function setCellStyle(Cell $cell, string $type = null)
    {
        if (self::CELL_TYPE_WEEKDAY === $type) {
            $cell->getStyle()->getFill()->getStartColor()->setRGB(self::CELL_COLOR_WEEKDAY);
            $cell->getStyle()->getFill()->setFillType(Fill::FILL_SOLID);
        }
        if (self::CELL_TYPE_PAIR_NUMBER === $type) {
            $cell->getStyle()->getFill()->getStartColor()->setRGB(self::CELL_COLOR_PAIR_DAY);
            $cell->getStyle()->getFill()->setFillType(Fill::FILL_SOLID);
        }

        $this->finalRow = $this->finalRow > $cell->getRow() ? $this->finalRow : $cell->getRow();
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function saveDocument()
    {
        $writer = new Xls($this->spreadsheet);

        $writer->save('schedule.xls');
    }
}
