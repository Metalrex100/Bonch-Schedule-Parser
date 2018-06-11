<?php

namespace App;

use Carbon\Carbon;
use Goutte\Client;
use Symfony\Component\DomCrawler\Crawler;
use Symfony\Component\HttpFoundation\Request;

class ScheduleParser
{
    /**
     * @var Client
     */
    private $client;

    /**
     * @var array
     */
    private $parameters = [
        'schet'    => '',
        'type_z'   => 4,
        'faculty'  => 56580,
        'kurs'     => 0,
        'group'    => 53262,
        'ok'       => 'Показать',
        'group_el' => 0,
    ];

    /**
     * ScheduleParser constructor.
     */
    public function __construct()
    {
        $this->client = new Client();
    }

    const SCHEDULE_URL = 'https://cabinet.sut.ru/raspisanie_all_new';

    /**
     * @return array
     */
    public function parse()
    {
        $this->getCurrentTerm();

        return $this->getSchedule();
    }

    protected function getCurrentTerm()
    {
        $response = $this->client->request(Request::METHOD_GET, self::SCHEDULE_URL);

        $this->parameters['schet'] = $response->filter('form option[selected="selected"]')->first()->attr('value');
    }

    protected function getCurrentCourse()
    {
        switch (Carbon::now()->year) {
            case 2018:
                $this->parameters['kurs'] = 2;
                break;
            case 2019:
                $this->parameters['kurs'] = 3;
                break;
            case 2020:
                $this->parameters['kurs'] = 4;
                break;
        }
    }

    /**
     * @return array
     */
    protected function getSchedule()
    {
        $response = $this->client->request(Request::METHOD_POST, self::SCHEDULE_URL, $this->parameters);
        $scheduleTable = $response->filter('table.simple-little-table')->first();

        return $this->parseScheduleTable($scheduleTable);
    }

    /**
     * @param Crawler $scheduleTable
     *
     * @return array
     */
    protected function parseScheduleTable(Crawler $scheduleTable)
    {
        $schedule = [];
        $weekNumber = 0;
        $lastWeekday = 0;

        $scheduleTable->filter('tr.pair')->each(function (Crawler $node) use (&$schedule, &$weekNumber, &$lastWeekday) {
            $weekday = (int) $node->attr('weekday');
            $pairInfo = explode(' ', $node->attr('pair'));
            $pairNumber = array_first($pairInfo);
            $pairTime = strtr(array_last($pairInfo), ['(' => '', ')' => '']);
            $pairDate = substr($node->filter('td')->first()->text(), 0, 10);
            $pairType = $node->filter('td span.type')->text();
            $pairSubject = $node->filter('td span.subect')->text();
            $pairTeacher = $node->filter('td span.teacher')->text();
            $pairAddress = $node->filter('td span.aud')->text();

            if ($lastWeekday > 1 && $weekday === 1) {
                ++$weekNumber;
            }

            $schedule[$weekNumber][$weekday][$pairDate][] = [
                'number'  => $pairNumber,
                'time'    => $pairTime,
                'type'    => $pairType,
                'subject' => $pairSubject,
                'teacher' => $pairTeacher,
                'address' => $pairAddress,
            ];

            $lastWeekday = $weekday;
        })
        ;

        return $schedule;
    }
}
