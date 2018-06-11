<?php

namespace App\Console\Commands;

use App\ScheduleExcelExporter;
use App\ScheduleParser;
use Illuminate\Console\Command;

class ParseSchedule extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'schedule:parse';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Parses university schedule';

    /**
     * @var ScheduleParser
     */
    private $scheduleParser;

    /**
     * @var ScheduleExcelExporter
     */
    private $scheduleExcelExporter;

    /**
     * Create a new command instance.
     *
     * @param ScheduleParser        $scheduleParser
     * @param ScheduleExcelExporter $scheduleExcelExporter
     */
    public function __construct(ScheduleParser $scheduleParser, ScheduleExcelExporter $scheduleExcelExporter)
    {
        parent::__construct();

        $this->scheduleParser = $scheduleParser;
        $this->scheduleExcelExporter = $scheduleExcelExporter;
    }

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $parsedSchedule = $this->scheduleParser->parse();
        $this->scheduleExcelExporter->export($parsedSchedule);
    }
}
