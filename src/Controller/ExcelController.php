<?php

namespace App\Controller;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Annotation\Route;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;

class ExcelController extends AbstractController
{
    #[Route('/excel', name: 'app_excel')]
    public function transformJsonToExcel(): Response
    {
        // Ruta de la carpeta que contiene los archivos JSON
        $jsonFolderPath = $this->getParameter('kernel.project_dir') . '/public/archivosJson';

        // Obtener la lista de archivos JSON en la carpeta
        $jsonFiles = glob($jsonFolderPath . '/*.json');


        // Ruta del archivo Excel predefinido
        $excelFilePath = $this->getParameter('kernel.project_dir') . '/public/archivosExcel/Modelo Excel.xlsx';

        // Cargar el archivo Excel existente
        $spreadsheet = IOFactory::load($excelFilePath);

        // Obtener la hoja de cálculo activa
        $sheet = $spreadsheet->getActiveSheet();

        // Variable de control para la celda
        $initialRow = 3; // Comienza en la fila 3

        // Recorrer los archivos JSON
        foreach ($jsonFiles as $jsonFile) {
            // Leer el contenido del archivo JSON
            $jsonData = file_get_contents($jsonFile);

            // Decodificar el JSON a un arreglo asociativo
            $data = json_decode($jsonData, true);

            // Asignar el nombre del archivo JSON que se procesa a la columna A
            $sheet->setCellValue('A' . $initialRow, basename($jsonFile));

            // Realizar las asignaciones correspondientes
            $this->assignDataToExcel($sheet, $data, $initialRow);

            // Incrementar la variable de control de la fila
            $initialRow++;
        }

        // Guardar el archivo Excel modificado
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($excelFilePath);

        // Devolver una respuesta con el enlace para descargar el archivo Excel actualizado
        return $this->file($excelFilePath);
    }

    //----------------------------------------------------------------------------------------------------------------------------------------------------------------//

    /**
     * Asigna los datos del archivo JSON a las celdas correspondientes en el archivo Excel.
     *
     * @param Worksheet $sheet La hoja de cálculo activa.
     * @param array $data Los datos del archivo JSON.
     * @param int $initialRow La fila inicial para realizar las asignaciones.
     */

    private function assignDataToExcel(Worksheet $sheet, array $data, int $initialRow): void
    {
        // 1-Busqueda de variables en ANSWER
        if (isset($data['processes']['GPT4 8K Azure']['result']['answer'])) {
            $answer = $data['processes']['GPT4 8K Azure']['result']['answer'];

            // Obtener el valor de fileName de la sección "Categorizador_GPT4 8K Azure"
            $fileName = $data['processes']['Categorizador_GPT4 8K Azure']['result']['fileName'];

            // Agregar el valor de fileName a la columna A
            $sheet->setCellValue('A' . $initialRow, $fileName);


            // Definir las variables de los campos del archivo JSON en la columna que corresponde al Excel
            $cellMapping = [
                'B' => 'fileSubType',
                'C' => 'Tipo',
                'D' => 'Tipo de sociedad',
                'G' => 'CIF',
            ];
            //----------------------------------------- Razon Social ---------------------------------------------//
            if (isset($answer['Razón Social'][0])) {
                $razonSocial = $answer['Razón Social'][0];
                $sheet->setCellValue('E' . $initialRow, $razonSocial['Denominación']);
                $sheet->setCellValue('F' . $initialRow, $razonSocial['Domicilio social']);
            }
            //-----------------------------------Inscripcion Registro Mercantil-------------------------------//
            if (isset($answer['Inscripción Reg. Mercantil'][0])) {
                $registroMercantil = $answer['Inscripción Reg. Mercantil'][0];
                $sheet->setCellValue('H' . $initialRow, $registroMercantil['Inscrito']);
                $sheet->setCellValue('I' . $initialRow, $registroMercantil['Hoja']);
                $sheet->setCellValue('J' . $initialRow, $registroMercantil['Tomo']);
                $sheet->setCellValue('K' . $initialRow, $registroMercantil['Inscripción']);
            }

            //---------------------------------------Notarios hasta 3----------------------------------------//
            if (isset($answer['Notario'][0])) {
                $notario1 = $answer['Notario'][0];
                $sheet->setCellValue('L' . $initialRow, $notario1['Nombre/Apellido']);
                $sheet->setCellValue('M' . $initialRow, $notario1['Num. protocolo']);
                $sheet->setCellValue('N' . $initialRow, $notario1['Fecha escritura']);
                $sheet->setCellValue('O' . $initialRow, $notario1['Localidad']);
                $sheet->setCellValue('P' . $initialRow, $notario1['Col. notarios']);
            }
            if (isset($answer['Notario'][1])) {
                $notario2 = $answer['Notario'][1];
                $sheet->setCellValue('Q' . $initialRow, $notario2['Nombre/Apellido']);
                $sheet->setCellValue('R' . $initialRow, $notario2['Num. protocolo']);
                $sheet->setCellValue('S' . $initialRow, $notario2['Fecha escritura']);
                $sheet->setCellValue('T' . $initialRow, $notario2['Localidad']);
                $sheet->setCellValue('U' . $initialRow, $notario2['Col. notarios']);
            }

            if (isset($answer['Notario'][2])) {
                $notario3 = $answer['Notario'][2];
                $sheet->setCellValue('V' . $initialRow, $notario3['Nombre/Apellido']);
                $sheet->setCellValue('W' . $initialRow, $notario3['Num. protocolo']);
                $sheet->setCellValue('X' . $initialRow, $notario3['Fecha escritura']);
                $sheet->setCellValue('Y' . $initialRow, $notario3['Localidad']);
                $sheet->setCellValue('Z' . $initialRow, $notario3['Col. notarios']);
            }
            //--------------------------------------Apoderados hasta 3---------------------------------//
            if (isset($answer['Apoderado'][0])) {
                $apoderado1 = $answer['Apoderado'][0];
                $sheet->setCellValue('AA' . $initialRow, $apoderado1['Nombres']);
                $sheet->setCellValue('AB' . $initialRow, $apoderado1['Apellidos']);
                $sheet->setCellValue('AC' . $initialRow, $apoderado1['Número DNI']);
                $sheet->setCellValue('AD' . $initialRow, $apoderado1['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('AE' . $initialRow, $apoderado1['Domicilio - Nombre']);
                $sheet->setCellValue('AF' . $initialRow, $apoderado1['Domicilio - Número']);
                $sheet->setCellValue('AG' . $initialRow, $apoderado1['Domicilio - Localidad']);
                $sheet->setCellValue('AH' . $initialRow, $apoderado1['Domicilio - Provincia']);
                $sheet->setCellValue('AI' . $initialRow, $apoderado1['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][1])) {
                $apoderado2 = $answer['Apoderado'][1];
                $sheet->setCellValue('AJ' . $initialRow, $apoderado2['Nombres']);
                $sheet->setCellValue('AK' . $initialRow, $apoderado2['Apellidos']);
                $sheet->setCellValue('AL' . $initialRow, $apoderado2['Número DNI']);
                $sheet->setCellValue('AM' . $initialRow, $apoderado2['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('AN' . $initialRow, $apoderado2['Domicilio - Nombre']);
                $sheet->setCellValue('AO' . $initialRow, $apoderado2['Domicilio - Número']);
                $sheet->setCellValue('AP' . $initialRow, $apoderado2['Domicilio - Localidad']);
                $sheet->setCellValue('AQ' . $initialRow, $apoderado2['Domicilio - Provincia']);
                $sheet->setCellValue('AR' . $initialRow, $apoderado2['Tipo de apoderamiento']);
            }

            if (isset($answer['Apoderado'][2])) {
                $apoderado3 = $answer['Apoderado'][2];
                $sheet->setCellValue('AS' . $initialRow, $apoderado3['Nombres']);
                $sheet->setCellValue('AT' . $initialRow, $apoderado3['Apellidos']);
                $sheet->setCellValue('AU' . $initialRow, $apoderado3['Número DNI']);
                $sheet->setCellValue('AV' . $initialRow, $apoderado3['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('AW' . $initialRow, $apoderado3['Domicilio - Nombre']);
                $sheet->setCellValue('AX' . $initialRow, $apoderado3['Domicilio - Número']);
                $sheet->setCellValue('AY' . $initialRow, $apoderado3['Domicilio - Localidad']);
                $sheet->setCellValue('AZ' . $initialRow, $apoderado3['Domicilio - Provincia']);
                $sheet->setCellValue('BA' . $initialRow, $apoderado3['Tipo de apoderamiento']);
            }
        }

        // Recorrer el mapeo y asignar los valores del archivo JSON a las celdas correspondientes
        foreach ($cellMapping as $column => $field) {
            $cell = $column . $initialRow;
            $value = isset($answer[$field]) ? $answer[$field] : '';
            $sheet->setCellValue($cell, $value);
        }
    }
}
