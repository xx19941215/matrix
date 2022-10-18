<?php

namespace App\Console\Commands;

use App\Model\Letter;
use App\Model\LetterDep;
use App\Model\LetterType;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Log;


class ExportMatrix extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'export:matrix {filePath}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Command description';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return mixed
     */
    public function handle()
    {
        $filepath = storage_path($this->argument("filePath"));

        $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
        $objReader->setReadDataOnly(true);

        $objPHPExcel = $objReader->load($filepath);

        $sheet = $objPHPExcel->getSheet(0);
//获取sheet页的名字：$sheetName = $sheet->getTitle();
        $highestRow = $sheet->getHighestRow(); // 取得总行数

        $data = [];
        $data2 = [];
        $j = [];

        for ($key = 2; $key <= $highestRow; $key++) {

            $jl = $sheet->getCellByColumnAndRow(0, $key)->getValue();//获取菌落
            $jy = $sheet->getCellByColumnAndRow(1, $key)->getValue();//获取基因
// if (is_bool(strpos($jl, 'D50'))) {
//     $j[] = $jy;
// }
            Log::info(var_export($jl, true));
            Log::info(var_export($jy, true));

            if (isset($data, $jy)) {
                $data[$jy][$jl] = 1;
                $data2[$jl][$jy] = 1;
            } else {
                $new = [];
                $new[$jl] = 1;
                $data[$jy] = $new;

                $new2 = [];
                $new2[$jy] = 1;
                $data2[$jl] = $new;
            }
        }


        $objPHPExcel = new \PHPExcel();
        $sheet = self::getPHPExcel($objPHPExcel, $data);

        $columns = collect($data)->keys()->filter(function ($v) {
            return is_bool(strpos($v, 'D50')) && $v;
        })->values();


        $data = collect($data)->filter(function ($v, $k) {
            return is_bool(strpos($k, 'D50')) && $k;
        });
//        $j = array_unique($j);
//        array_pop($j);

        $gene = collect($data)->keys()->values();

//横坐标
        $alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

        $c = [];

        foreach ($alphabet as $d) {
            array_push($c, $d);
        }

        $three = [];

        foreach ($alphabet as $a) {
            foreach ($alphabet as $b) {
                array_push($three, $a . $b);
                array_push($c, $a . $b);
            }
        }

        foreach ($alphabet as $a) {
            foreach ($three as $t) {
                array_push($c, $a . $t);
            }
        }

        $alphabet = array_slice($c, 1, count($columns));

        $row = 1;
        $g_index = 0;
// dd($data2['0807D0.fasta'], $data['iroN_27']);
        foreach ($data2 as $k => $v) {
            $row++;
            $sheet->setCellValue("A{$row}", $k);
            $g_index++;

// dd();
            foreach ($gene as $ind => $g) {
// dd($gene, $data[$g]);
                $value = 0;
                if (array_key_exists($k, $data[$g])) {
                    $value = 1;
                }
// $col_index = $ind+1;
// dd($alphabet);
                $sheet->setCellValue("{$alphabet[$ind]}{$row}", $value);

// foreach ($alphabet as $a) {
//     //$g基因和坐标$a是对应的
//     // $sheet->setCellValue("$a{$row}",1);
//     // dd(array_key_exists($k, $data[$g]) ? 1 : 0);
//     $value = 0;
//     if (array_key_exists($k, $data[$g])) {
//         $value = 1;
//     }
//     $sheet->setCellValue("$a{$row}", $value);
// }
            }
        }


        $path = self::get_tmp_path() . sprintf("/j_%s.xlsx", date('Y-m-d-H-i-s'));
        (new \PHPExcel_Writer_Excel2007($objPHPExcel))->save($path);
    }


    private static function get_tmp_path()
    {
        $path = storage_path('app/tmp/command-excel');

        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }

        return $path;
    }

    protected static function getPHPExcel(\PHPExcel $objPHPExcel, $data)
    {

// $objPHPExcel->getDefaultStyle()->applyFromArray([
//     'font' => ['size' => 12, 'name' => '宋体'],
//     'alignment' => [
//         'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
//         'vertical' => \PHPExcel_Style_Alignment::VERTICAL_CENTER,
//         'wrap' => true,
//     ],
// ]);

        $sheet = $objPHPExcel->getActiveSheet();


        $columns = collect($data)->keys()->filter(function ($v) {
            return is_bool(strpos($v, 'D50')) && $v;
        })->values();


        $alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

        $c = [];

        foreach ($alphabet as $d) {
            array_push($c, $d);
        }

        $three = [];
        foreach ($alphabet as $a) {
            foreach ($alphabet as $b) {
                array_push($three, $a . $b);
                array_push($c, $a . $b);
            }
        }

        foreach ($alphabet as $a) {
            foreach ($three as $t) {
                array_push($c, $a . $t);
            }
        }


        $alphabet = array_slice($c, 1, count($columns));
        for ($i = 0; $i < count($columns); $i++) {
            $sheet->setCellValue($alphabet[$i] . '1', $columns[$i]);
        }

        return $sheet;
    }
}