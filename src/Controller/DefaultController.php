<?php

namespace App\Controller;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use Symfony\Component\DependencyInjection\ParameterBag\ParameterBagInterface;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;
use Symfony\Component\Routing\Attribute\Route;

class DefaultController extends AbstractController
{
    #[Route('/', name: 'presupuesto')]
    public function presupuesto(ParameterBagInterface $params)
    {
        // Create the spreadsheet
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getDefaultStyle()->applyFromArray([
            'font' => [
                'size' => 12,
                'name' => 'Arial'
            ]
        ]);

        // Sheet title
        $sheet = $spreadsheet->getActiveSheet()->setTitle('Portada');

        // Borders
        $topBorderThinStyle = ['borders' => ['top' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']]]];
        $rightBorderThinStyle = ['borders' => ['right' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']]]];
        $bottomBorderThinStyle = ['borders' => ['bottom' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']]]];
        $leftBorderThinStyle = ['borders' => ['left' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']]]];
        $bordersThinStyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];
        $bordersThickStyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];
        // Merge
        $sheet->mergeCells([4, 15, 8, 18]);

        // Borders
        $sheet->getStyle([2, 2, 10, 2])->applyFromArray($topBorderThinStyle);
        $sheet->getStyle([10, 2, 10, 37])->applyFromArray($rightBorderThinStyle);
        $sheet->getStyle([2, 37, 10, 37])->applyFromArray($bottomBorderThinStyle);
        $sheet->getStyle([2, 2, 2, 37])->applyFromArray($leftBorderThinStyle);
        $sheet->getStyle([2, 2, 10, 37])->getFill()->setFillType('solid')->getStartColor()->setRGB('ffffff');
        $sheet->getStyle([4, 15, 8, 18])->applyFromArray($bordersThickStyle);

        // Width
        $sheet->getColumnDimensionByColumn(2)->setWidth(12);
        $sheet->getColumnDimensionByColumn(3)->setWidth(12);
        $sheet->getColumnDimensionByColumn(4)->setWidth(12);
        $sheet->getColumnDimensionByColumn(5)->setWidth(12);
        $sheet->getColumnDimensionByColumn(6)->setWidth(12);
        $sheet->getColumnDimensionByColumn(7)->setWidth(12);
        $sheet->getColumnDimensionByColumn(8)->setWidth(12);
        $sheet->getColumnDimensionByColumn(9)->setWidth(12);
        $sheet->getColumnDimensionByColumn(10)->setWidth(12);

        // General style
        $sheet->getStyle([4, 15])->getAlignment()->setHorizontal('center');
        $sheet->getStyle([4, 15])->getAlignment()->setVertical('center');
        $sheet->getStyle([4, 15])->getFill()->setFillType('solid')->getStartColor()->setRGB('d8edf5');
        $sheet->getStyle([4, 15])->getFont()->setSize(24);
        $sheet->getStyle([4, 15])->getFont()->setBold(true);

        $sheet->getStyle([4, 15])->getFont()->setSize(24);

        $sheet->getStyle([5, 20])->getFont()->setSize(9);
        $sheet->getStyle([5, 22])->getFont()->setSize(9);

        // Values
        $sheet
            ->setCellValue([4, 15], 'Presupuesto')
            ->setCellValue([4, 20], 'Cliente')
            ->setCellValue([5, 20], 'Mercadona')
            ->setCellValue([4, 22], 'Fecha')
            ->setCellValue([5, 22], '10/04/2024');

            // Create another sheet
        $spreadsheet->createSheet(1);
        $spreadsheet->setActiveSheetIndex(1);
        $sheet = $spreadsheet->getActiveSheet()->setTitle('Items');

        $spreadsheet->getDefaultStyle()->applyFromArray([
            'font' => [
                'size' => 10,
                'name' => 'Arial'
            ]
        ]);

        // Set height and width
        $sheet->getRowDimension(6)->setRowHeight(22);
        $sheet->getRowDimension(7)->setRowHeight(22);
        $sheet->getRowDimension(8)->setRowHeight(22);
        $sheet->getRowDimension(9)->setRowHeight(22);
        $sheet->getColumnDimensionByColumn(1)->setWidth(15);
        $sheet->getColumnDimensionByColumn(2)->setWidth(50);
        $sheet->getColumnDimensionByColumn(3)->setWidth(12);
        $sheet->getColumnDimensionByColumn(4)->setWidth(15);
        $sheet->getColumnDimensionByColumn(5)->setWidth(15);

        // Table header
        $sheet->getStyle([1, 6, 5, 6])->applyFromArray($bordersThinStyle);
        $sheet->getStyle([1, 6, 5, 6])->getAlignment()->setVertical('center');
        $sheet->getStyle([1, 6, 5, 6])->getAlignment()->setHorizontal('center');
        $sheet->getStyle([1, 6, 5, 6])->getFill()->setFillType('solid')->getStartColor()->setRGB('d8edf5');
        $sheet->getStyle([1, 6, 5, 6])->getFont()->setBold(true);

        $sheet
            ->setCellValue([1, 6], 'Pos.')
            ->setCellValue([2, 6], 'Concepto')
            ->setCellValue([3, 6], 'Cantidad')
            ->setCellValue([4, 6], 'Precio unidad')
            ->setCellValue([5, 6], 'Total');

        // Table content
        $sheet->getStyle([5, 7, 5, 9])->applyFromArray($rightBorderThinStyle);
        $sheet->getStyle([1, 9, 5, 9])->applyFromArray($bottomBorderThinStyle);
        $sheet->getStyle([1, 7, 5, 9])->getAlignment()->setVertical('center');
        $sheet->getStyle([1, 7, 1, 9])->getAlignment()->setHorizontal('center');
        $sheet->getStyle([3, 7, 5, 9])->getAlignment()->setHorizontal('center');

        $sheet
            ->setCellValue([1, 7], '1')
            ->setCellValue([2, 7], 'Servicios de desarrollo de software')
            ->setCellValue([3, 7], '2')
            ->setCellValue([4, 7], '600 €')
            ->setCellValue([5, 7], '1200 €')
            ->setCellValue([1, 8], '2')
            ->setCellValue([2, 8], 'Implementación de dominio')
            ->setCellValue([3, 8], '2')
            ->setCellValue([4, 8], '20 €')
            ->setCellValue([5, 8], '40 €')
            ->setCellValue([1, 9], '3')
            ->setCellValue([2, 9], 'Contratación de hosting')
            ->setCellValue([3, 9], '1')
            ->setCellValue([4, 9], '15 €')
            ->setCellValue([5, 9], '15 €');

        // Observations
        $sheet->getStyle([2, 14])->applyFromArray($topBorderThinStyle);
        $sheet->getStyle([2, 14, 2, 17])->applyFromArray($rightBorderThinStyle);
        $sheet->getStyle([2, 17])->applyFromArray($bottomBorderThinStyle);
        $sheet->getStyle([2, 14, 2, 17])->applyFromArray($leftBorderThinStyle);
        $sheet->getStyle([2, 14, 2, 17])->getFill()->setFillType('solid')->getStartColor()->setRGB('ffffff');
        $sheet->getStyle([2, 14])->getFont()->setBold(true);

        $sheet->setCellValue([2, 14], 'Observaciones:');

        // Observations
        $sheet->getStyle([4, 14, 5, 14])->applyFromArray($topBorderThinStyle);
        $sheet->getStyle([5, 14, 5, 17])->applyFromArray($rightBorderThinStyle);
        $sheet->getStyle([4, 17, 5, 17])->applyFromArray($bottomBorderThinStyle);
        $sheet->getStyle([4, 14, 4, 17])->applyFromArray($leftBorderThinStyle);
        $sheet->getStyle([4, 14, 5, 17])->getFill()->setFillType('solid')->getStartColor()->setRGB('ffffff');
        $sheet->getStyle([4, 17, 5, 17])->getFont()->setBold(true);
        $sheet->getStyle([4, 14, 5, 17])->getAlignment()->setHorizontal('center');

        $sheet->setCellValue([4, 14], 'Subtotal');
        $sheet->setCellValue([5, 14], '1255 €');
        $sheet->setCellValue([4, 15], 'IVA (21%)');
        $sheet->setCellValue([5, 15], '263.55 €');
        $sheet->setCellValue([4, 17], 'TOTAL');
        $sheet->setCellValue([5, 17], '1518.55 €');

        $spreadsheet->setActiveSheetIndex(0);

        $writer = IOFactory::createWriter($spreadsheet, "Xlsx");
        $path = "{$params->get('kernel.project_dir')}/var/xls/";
        $fileName = "presupuesto.xlsx";

        $writer->save("{$path}{$fileName}");

        return $this->file("{$path}{$fileName}", $fileName, ResponseHeaderBag::DISPOSITION_ATTACHMENT)->deleteFileAfterSend();
    }
}