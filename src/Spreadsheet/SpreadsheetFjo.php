<?php

namespace fjourneau\Spreadsheet;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use App\Objects\fjoUtilitiesObject as Utilities;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use Slim\Http\Response;

/**
 * Classe améliorée pour sites FJO pour génération extract XLS
 *
 * @author    fJourneau
 */
class SpreadsheetFjo extends Spreadsheet
{

    /**
     * Lien entre n° et nom de colonne (1 => A, 2 => B, ..., 27 => AA...)
     * @var array Tableau renseigné dans $this->_fillCollArray()
     */
    private $colArray = [];

    /**
     * N° de ligne courante du cursor
     * 
     * @var int N° de ligne
     */
    protected $row = 0;

    /**
     * N° de ligne initial du cursor
     *
     * @var int N° de ligne
     */
    protected $StartRow = 0;

    /**
     * N° de colonne courant du cursor (1 => A, 2 => B, etc...)
     * 
     * @var int N° de colonne
     */
    protected $col = 0;

    /**
     * N° de colonne initial du cursor (1 => A, 2 => B, etc...)
     *
     * @var int N° de colonne
     */
    protected $startCol = 0;

    public function __construct()
    {
        parent::__construct();

        $this->_fillCollArray();
    }

    /**
     * Rajout de propriétés sur le spreadsheet 
     *
     * @param  array [creator, title, lastModifier, description, keywords, category]
     * @return void
     */
    public function setSpreadsheetProperties(array $options): void
    {
        $this->getProperties()->setCreator($options['creator'] ?? 'Florian JOURNEAU');
        $this->getProperties()->setTitle($options['title'] ?? '');

        if (isset($options['lastModifier'])) {
            $this->getProperties()->setLastModifiedBy($options['lastModifier']);
        }
        if (isset($options['description'])) {
            $this->getProperties()->setLastModifiedBy($options['description']);
        }
        if (isset($options['keywords'])) {
            $this->getProperties()->setLastModifiedBy($options['keywords']);
        }
        if (isset($options['category'])) {
            $this->getProperties()->setLastModifiedBy($options['category']);
        }
    }

    /**
     * Initialise la position du curseur pour créer un tableau
     *
     * @param  int $col
     * @param  int $row
     * @return void
     */
    public function initCusor(int $col, int $row)
    {
        $this->col = $this->startCol = $col;
        $this->row = $this->StartRow = $row;
    }

    /**
     * Mettre une valeur dans une cellule où le curseur se trouve
     *
     * @param  mixed $val Valeur
     * @return void
     */
    public function setCursorValue(string $val)
    {
        $cellName = $this->_getCellNameFromCursor();

        $this->getActiveSheet()->setCellValue($cellName, $val);
    }

    /**
     * Mettre un prix dans une cellule où le curseur se trouve
     *
     * @param  float $val Valeur numérique (prix)
     * @param  string $curr 'EUR' ou 'USD'
     * @return void
     */
    public function setCursorValueCurrency(float $val, $curr = 'EUR')
    {
        $cellName = $this->_getCellNameFromCursor();
        if ($curr == 'EUR') {
            $this->getActiveSheet()->getStyle($cellName)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
        } elseif ($curr == 'USD') {
            $this->getActiveSheet()->getStyle($cellName)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
        }
        $this->getActiveSheet()->setCellValue($cellName, $val);
    }

    /**
     * Déplace le curseur à droite sur la même ligne
     *
     * @return void
     */
    public function moveCursor(): SpreadsheetFjo
    {
        $this->col++;
        return $this;
    }

    /**
     * Faire passer le curseur sur une nouvelle ligne
     *
     * @return void
     */
    public function newLineCursor()
    {
        $this->col = $this->startCol;
        $this->row++;

        return $this;
    }


    public function setHeaderStyle()
    {
        $range = $this->colArray[$this->startCol] . $this->StartRow . ':' . $this->_getCellNameFromCursor();

        $this->getActiveSheet()->getStyle($range)->applyFromArray(
            [
                'font' => [
                    'bold' => true,
                ],
                'borders' => [
                    'top' => ['borderStyle' => Border::BORDER_THIN],
                    'bottom' => ['borderStyle' => Border::BORDER_THIN],
                    'left' => ['borderStyle' => Border::BORDER_THIN],
                    'right' => ['borderStyle' => Border::BORDER_THIN],
                ],
                'fill' => [
                    'fillType' => Fill::FILL_GRADIENT_LINEAR,
                    'rotation' => 90,
                    'startColor' => ['argb' => 'FFA0A0A0'],
                    'endColor' => ['argb' => 'FFFFFFFF'],
                ],
            ]
        );
    }

    /**
     * Défini l'onglet actif et lui assigne un titre.
     *
     * @param  string $title
     * @param  int $index
     * @return void
     */
    public function setTabTitle(string $title, int $index = 0)
    {
        $this->setActiveSheetIndex($index);
        $this->getActiveSheet()->setTitle($title);
    }

    public function setSheetTitle(string $title, string $range = 'A1:I1'): void
    {
        $this->getActiveSheet()->mergeCells($range);
        $this->getActiveSheet()->getStyle('A1')->getFont()
            ->setSize(20)
            ->setBold(true)
            ->getColor()->setARGB(Color::COLOR_RED);

        $this->getActiveSheet()->setCellValue('A1', $title);
    }

    /**
     * Génère le fichier XLSX à télécharger ou sauver sur le serveur
     *
     * @param  Response $response (Slim Response)
     * @param  string $filename Nom du fichier à télécharger ou endroit où sauvegarder sur le serveur
     * @param  mixed $download TRUE pour téléchargement ou FALSE pour sauver sur le serveur
     * @return Response
     */
    public function generateExcelFile(?Response $response, $filename = 'file.xlsx', bool $download = true): Response
    {
        if ($download) {
            $responseXls = $response->withHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                ->withHeader('Content-Disposition', 'attachment;filename="' . $filename . '"')
                ->withHeader('Cache-Control', 'max-age=0')
                ->withHeader('Expires', 'Mon, 26 Jul 1997 05:00:00 GMT')  // date in the past
                ->withHeader('Last-Modified', gmdate('D, d M Y H:i:s') . ' GMT')
                ->withHeader('Cache-Control', 'cache, must-revalidate')  // HTTP/1.1
                ->withHeader('Pragma', 'private');   // HTTP/1.0

            $writer = IOFactory::createWriter($this, 'Xlsx');

            ob_start();
            $writer->save('php://output');
            $xls_content = ob_get_clean();

            return $responseXls->write($xls_content);
        } else {
            $writer = IOFactory::createWriter($this, 'Xlsx');
            $writer->save($filename);
            return $response->write("Fichier $filename généré à la racine du site.");
        }
    }

    /**
     * Génère le fichier ODS à télécharger ou sauver sur le serveur
     *
     * @param  Response $response (Slim Response)
     * @param  string $filename Nom du fichier à télécharger ou endroit où sauvegarder sur le serveur
     * @param  mixed $download TRUE pour téléchargement ou FALSE pour sauver sur le serveur
     * @return Response
     */
    public function generateOdsFile(?Response $response, $filename = 'file.ods', bool $download = true): Response
    {
        if ($download) {
            $responseXls = $response->withHeader('Content-Type', 'application/vnd.oasis.opendocument.text')
                ->withHeader('Content-Disposition', 'attachment;filename="' . $filename . '"')
                ->withHeader('Cache-Control', 'max-age=0')
                ->withHeader('Expires', 'Mon, 26 Jul 1997 05:00:00 GMT')  // date in the past
                ->withHeader('Last-Modified', gmdate('D, d M Y H:i:s') . ' GMT')
                ->withHeader('Cache-Control', 'cache, must-revalidate')  // HTTP/1.1
                ->withHeader('Pragma', 'private');   // HTTP/1.0

            $writer = IOFactory::createWriter($this, 'Ods');

            ob_start();
            $writer->save('php://output');
            $xls_content = ob_get_clean();

            return $responseXls->write($xls_content);
        } else {
            $writer = IOFactory::createWriter($this, 'Ods');
            $writer->save($filename);
            return $response->write("Fichier $filename généré à la racine du site.");
        }
    }

    /**
     * Génère le fichier ODS à télécharger ou sauver sur le serveur
     *
     * @param  Response $response (Slim Response)
     * @param  string $filename Nom du fichier à télécharger ou endroit où sauvegarder sur le serveur
     * @param  mixed $download TRUE pour téléchargement ou FALSE pour sauver sur le serveur
     * @return Response
     */
    public function generatePdfFile(?Response $response, $filename = 'file.ods', bool $download = true): Response
    {
        if ($download) {
            $responseXls = $response->withHeader('Content-Type', 'application/pdf')
                ->withHeader('Content-Disposition', 'attachment;filename="' . $filename . '"')
                ->withHeader('Cache-Control', 'max-age=0')
                ->withHeader('Expires', 'Mon, 26 Jul 1997 05:00:00 GMT')  // date in the past
                ->withHeader('Last-Modified', gmdate('D, d M Y H:i:s') . ' GMT')
                ->withHeader('Cache-Control', 'cache, must-revalidate')  // HTTP/1.1
                ->withHeader('Pragma', 'private');   // HTTP/1.0

            $writer = IOFactory::createWriter($this, 'Dompdf');

            ob_start();
            $writer->save('php://output');
            $xls_content = ob_get_clean();

            return $responseXls->write($xls_content);
        } else {
            $writer = IOFactory::createWriter($this, 'Dompdf');
            $writer->save($filename);
            return $response->write("Fichier $filename généré à la racine du site.");
        }
    }


    public function setColumnsWidth(array $cols)
    {
        foreach ($cols as  $colName => $width) {
            if ($width == 'auto') {
                $this->getActiveSheet()->getColumnDimension($colName)->setAutoSize(true);
            } else {
                $this->getActiveSheet()->getColumnDimension($colName)->setWidth($width);
            }
        }
    }


    protected function _getCellNameFromCursor(): string
    {
        return $this->colArray[$this->col] . $this->row;
    }

    protected function _fillCollArray(): void
    {
        $this->colArray = [
            1 => 'A',
            2 => 'B',
            3 => 'C',
            4 => 'D',
            5 => 'E',
            6 => 'F',
            7 => 'G',
            8 => 'H',
            9 => 'I',
            10 => 'J',
            11 => 'K',
            12 => 'L',
            13 => 'M',
            14 => 'N',
            15 => 'O',
            16 => 'P',
            17 => 'Q',
            18 => 'R',
            19 => 'S',
            20 => 'T',
            21 => 'U',
            22 => 'V',
            23 => 'W',
            24 => 'X',
            25 => 'Y',
            26 => 'Z',
            27 => 'AA',
            28 => 'AB',
            29 => 'AC',
            30 => 'AD',
            31 => 'AE',
            32 => 'AF',
            33 => 'AG',
            34 => 'AH',
            35 => 'AI',
            36 => 'AJ',
            37 => 'AK',
            38 => 'AL',
            39 => 'AM',
            40 => 'AN',
            41 => 'AO',
            42 => 'AP',
            43 => 'AQ',
            44 => 'AR',
            45 => 'AS',
            46 => 'AT',
            47 => 'AU',
            48 => 'AV',
            49 => 'AW',
            50 => 'AX',
            51 => 'AY',
            52 => 'AZ',
            53 => 'BA',
            54 => 'BB',
            55 => 'BC',
            56 => 'BD',
            57 => 'BE',
            58 => 'BF',
            59 => 'BG',
            60 => 'BH',
            61 => 'BI',
            62 => 'BJ',
            63 => 'BK',
            64 => 'BL',
            65 => 'BM',
            66 => 'BN',
            67 => 'BO',
            68 => 'BP',
            69 => 'BQ',
            70 => 'BR',
            71 => 'BS',
            72 => 'BT',
            73 => 'BU',
            74 => 'BV',
            75 => 'BW',
            76 => 'BX',
            77 => 'BY',
            78 => 'BZ',
            79 => 'CA',
            80 => 'CB',
            81 => 'CC',
            82 => 'CD',
            83 => 'CE',
            84 => 'CF',
            85 => 'CG',
            86 => 'CH',
            87 => 'CI',
            88 => 'CJ',
            89 => 'CK',
            90 => 'CL',
            91 => 'CM',
            92 => 'CN',
            93 => 'CO',
            94 => 'CP',
            95 => 'CQ',
            96 => 'CR',
            97 => 'CS',
            98 => 'CT',
            99 => 'CU',
            100 => 'CV',
            101 => 'CW',
            102 => 'CX',
            103 => 'CY',
            104 => 'CZ'
        ];
    }
}
