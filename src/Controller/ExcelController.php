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
            // Recorrer el mapeo y asignar los valores del archivo JSON a las celdas correspondientes
            foreach ($cellMapping as $column => $field) {
                $cell = $column . $initialRow;
                $value = isset($answer[$field]) ? $answer[$field] : '';
                $sheet->setCellValue($cell, $value);
            }
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
            //---------------------------------------Notario----------------------------------------//
            if (isset($answer['Notario'][0])) {
                $notario1 = $answer['Notario'][0];
                $sheet->setCellValue('L' . $initialRow, $notario1['Nombre/Apellido']);
                $sheet->setCellValue('M' . $initialRow, $notario1['Num. protocolo']);
                $sheet->setCellValue('N' . $initialRow, $notario1['Fecha escritura']);
                $sheet->setCellValue('O' . $initialRow, $notario1['Localidad']);
                $sheet->setCellValue('P' . $initialRow, $notario1['Col. notarios']);
            }
            //--------------------------------------Apoderados hasta 8---------------------------------//
            if (isset($answer['Apoderado'][0])) {
                $apoderado1 = $answer['Apoderado'][0];
                $sheet->setCellValue('Q' . $initialRow, $apoderado1['Nombres']);
                $sheet->setCellValue('R' . $initialRow, $apoderado1['Apellidos']);
                $sheet->setCellValue('R' . $initialRow, $apoderado1['Número DNI']);
                $sheet->setCellValue('T' . $initialRow, $apoderado1['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('U' . $initialRow, $apoderado1['Domicilio - Nombre']);
                $sheet->setCellValue('V' . $initialRow, $apoderado1['Domicilio - Número']);
                $sheet->setCellValue('W' . $initialRow, $apoderado1['Domicilio - Localidad']);
                $sheet->setCellValue('X' . $initialRow, $apoderado1['Domicilio - Provincia']);
                $sheet->setCellValue('Y' . $initialRow, $apoderado1['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][1])) {
                $apoderado2 = $answer['Apoderado'][1];
                $sheet->setCellValue('Z' . $initialRow, $apoderado2['Nombres']);
                $sheet->setCellValue('AA' . $initialRow, $apoderado2['Apellidos']);
                $sheet->setCellValue('AB' . $initialRow, $apoderado2['Número DNI']);
                $sheet->setCellValue('AC' . $initialRow, $apoderado2['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('AD' . $initialRow, $apoderado2['Domicilio - Nombre']);
                $sheet->setCellValue('AE' . $initialRow, $apoderado2['Domicilio - Número']);
                $sheet->setCellValue('AF' . $initialRow, $apoderado2['Domicilio - Localidad']);
                $sheet->setCellValue('AG' . $initialRow, $apoderado2['Domicilio - Provincia']);
                $sheet->setCellValue('AH' . $initialRow, $apoderado2['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][2])) {
                $apoderado3 = $answer['Apoderado'][2];
                $sheet->setCellValue('AI' . $initialRow, $apoderado3['Nombres']);
                $sheet->setCellValue('AJ' . $initialRow, $apoderado3['Apellidos']);
                $sheet->setCellValue('AK' . $initialRow, $apoderado3['Número DNI']);
                $sheet->setCellValue('AL' . $initialRow, $apoderado3['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('AM' . $initialRow, $apoderado3['Domicilio - Nombre']);
                $sheet->setCellValue('AN' . $initialRow, $apoderado3['Domicilio - Número']);
                $sheet->setCellValue('AO' . $initialRow, $apoderado3['Domicilio - Localidad']);
                $sheet->setCellValue('AP' . $initialRow, $apoderado3['Domicilio - Provincia']);
                $sheet->setCellValue('AQ' . $initialRow, $apoderado3['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][3])) {
                $apoderado4 = $answer['Apoderado'][3];
                $sheet->setCellValue('AR' . $initialRow, $apoderado4['Nombres']);
                $sheet->setCellValue('AS' . $initialRow, $apoderado4['Apellidos']);
                $sheet->setCellValue('AT' . $initialRow, $apoderado4['Número DNI']);
                $sheet->setCellValue('AU' . $initialRow, $apoderado4['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('AV' . $initialRow, $apoderado4['Domicilio - Nombre']);
                $sheet->setCellValue('AW' . $initialRow, $apoderado4['Domicilio - Número']);
                $sheet->setCellValue('AX' . $initialRow, $apoderado4['Domicilio - Localidad']);
                $sheet->setCellValue('AY' . $initialRow, $apoderado4['Domicilio - Provincia']);
                $sheet->setCellValue('AZ' . $initialRow, $apoderado4['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][4])) {
                $apoderado5 = $answer['Apoderado'][4];
                $sheet->setCellValue('BA' . $initialRow, $apoderado5['Nombres']);
                $sheet->setCellValue('BB' . $initialRow, $apoderado5['Apellidos']);
                $sheet->setCellValue('BC' . $initialRow, $apoderado5['Número DNI']);
                $sheet->setCellValue('BD' . $initialRow, $apoderado5['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('BE' . $initialRow, $apoderado5['Domicilio - Nombre']);
                $sheet->setCellValue('BF' . $initialRow, $apoderado5['Domicilio - Número']);
                $sheet->setCellValue('BG' . $initialRow, $apoderado5['Domicilio - Localidad']);
                $sheet->setCellValue('BH' . $initialRow, $apoderado5['Domicilio - Provincia']);
                $sheet->setCellValue('BI' . $initialRow, $apoderado5['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][5])) {
                $apoderado6 = $answer['Apoderado'][5];
                $sheet->setCellValue('BJ' . $initialRow, $apoderado6['Nombres']);
                $sheet->setCellValue('BK' . $initialRow, $apoderado6['Apellidos']);
                $sheet->setCellValue('BL' . $initialRow, $apoderado6['Número DNI']);
                $sheet->setCellValue('BM' . $initialRow, $apoderado6['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('BN' . $initialRow, $apoderado6['Domicilio - Nombre']);
                $sheet->setCellValue('BO' . $initialRow, $apoderado6['Domicilio - Número']);
                $sheet->setCellValue('BP' . $initialRow, $apoderado6['Domicilio - Localidad']);
                $sheet->setCellValue('BQ' . $initialRow, $apoderado6['Domicilio - Provincia']);
                $sheet->setCellValue('BR' . $initialRow, $apoderado6['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][6])) {
                $apoderado7 = $answer['Apoderado'][6];
                $sheet->setCellValue('BS' . $initialRow, $apoderado7['Nombres']);
                $sheet->setCellValue('BT' . $initialRow, $apoderado7['Apellidos']);
                $sheet->setCellValue('BU' . $initialRow, $apoderado7['Número DNI']);
                $sheet->setCellValue('BV' . $initialRow, $apoderado7['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('BW' . $initialRow, $apoderado7['Domicilio - Nombre']);
                $sheet->setCellValue('BX' . $initialRow, $apoderado7['Domicilio - Número']);
                $sheet->setCellValue('BY' . $initialRow, $apoderado7['Domicilio - Localidad']);
                $sheet->setCellValue('BZ' . $initialRow, $apoderado7['Domicilio - Provincia']);
                $sheet->setCellValue('CA' . $initialRow, $apoderado7['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][7])) {
                $apoderado8 = $answer['Apoderado'][7];
                $sheet->setCellValue('CB' . $initialRow, $apoderado8['Nombres']);
                $sheet->setCellValue('CC' . $initialRow, $apoderado8['Apellidos']);
                $sheet->setCellValue('CD' . $initialRow, $apoderado8['Número DNI']);
                $sheet->setCellValue('CE' . $initialRow, $apoderado8['Domicilio - Tipo de Vía']);
                $sheet->setCellValue('CF' . $initialRow, $apoderado8['Domicilio - Nombre']);
                $sheet->setCellValue('CG' . $initialRow, $apoderado8['Domicilio - Número']);
                $sheet->setCellValue('CH' . $initialRow, $apoderado8['Domicilio - Localidad']);
                $sheet->setCellValue('CI' . $initialRow, $apoderado8['Domicilio - Provincia']);
                $sheet->setCellValue('CJ' . $initialRow, $apoderado8['Tipo de apoderamiento']);
            }
        }
    }
}
