<?php

namespace App\Controller;

use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Annotation\Route;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;

class EscriturasController extends AbstractController
{
    #[Route('/escrituras', name: 'app_escrituras')]
    public function transformJsonToExcel(): Response
    {
        // Ruta de la carpeta que contiene los archivos JSON
        $jsonFilePath = $this->getParameter('kernel.project_dir') . '/public/2-para ver/output.json';

        // Leer el contenido del archivo JSON
        $jsonData = file_get_contents($jsonFilePath);

        // Decodificar el JSON a un arreglo asociativo
        $json = json_decode($jsonData, true);

        // Ruta del archivo Excel predefinido
        $excelFilePath = $this->getParameter('kernel.project_dir') . '/public/archivosExcel/Escrituras Apoderamiento.xlsx';

        // Cargar el archivo Excel existente
        $spreadsheet = IOFactory::load($excelFilePath);

        // Obtener la hoja de cálculo activa
        $sheet = $spreadsheet->getActiveSheet();

        // Variable de control para la celda
        $initialRow = 3; // Comienza en la fila 3

        // Realizar las asignaciones correspondientes
        $this->assignDataToExcel($sheet, $json, $initialRow);

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
     * @param array $json Los datos del archivo JSON.
     * @param int $initialRow La fila inicial para realizar las asignaciones.
     */

    private function assignDataToExcel(Worksheet $sheet, array $json, int $initialRow): void
    {
        foreach ($json as $fileName => $data) {
            // Escribe el nombre del archivo en la columna A
            $sheet->setCellValue('A' . $initialRow, $fileName);

            // Define el mapeo de columnas para los otros campos
            $cellMapping = [
                'B' => 'fileSubType',
                'C' => 'Tipo',
                'D' => 'Tipo de sociedad',
                'G' => 'CIF',
            ];

            // Recorrer el mapeo y asignar los valores a las celdas de Excel
            foreach ($cellMapping as $column => $field) {
                $cell = $column . $initialRow;
                $value = isset($data[$field]) ? $data[$field] : '';
                $sheet->setCellValue($cell, is_string($value) ? $value : '');
            }
            //-----------------------------------------------------Apoderados hasta 8 ---------------------------------//
            if (isset($data['Apoderado'])) {

                $apoderados = $data['Apoderado'];

                //------------------------------------------ Asignar datos del primer apoderado
                if (isset($apoderados[0])) {
                    $apoderado1 = $apoderados[0];
                    $sheet->setCellValue('Q' . $initialRow, isset($apoderado1['Nombres']) ? $apoderado1['Nombres'] : '');
                    $sheet->setCellValue('R' . $initialRow, isset($apoderado1['Apellidos']) ? $apoderado1['Apellidos'] : '');
                    $sheet->setCellValue('S' . $initialRow, isset($apoderado1['Número DNI']) ? $apoderado1['Número DNI'] : '');
                    $sheet->setCellValue('Y' . $initialRow, isset($apoderado1['Tipo de apoderamiento']) ? $apoderado1['Tipo de apoderamiento'] : '');
                    // Verificar la estructura del domicilio del primer apoderado
                    if (isset($apoderado1['Domicilio'])) {
                        // Si la información del domicilio está en un objeto 'Domicilio'
                        $domicilio = $apoderado1['Domicilio'];
                        $sheet->setCellValue('T' . $initialRow, isset($domicilio['Tipo de Vía']) ? $domicilio['Tipo de Vía'] : '');
                        $sheet->setCellValue('U' . $initialRow, isset($domicilio['Nombre']) ? $domicilio['Nombre'] : '');
                        $sheet->setCellValue('V' . $initialRow, isset($domicilio['Número']) ? $domicilio['Número'] : '');
                        $sheet->setCellValue('W' . $initialRow, isset($domicilio['Localidad']) ? $domicilio['Localidad'] : '');
                        $sheet->setCellValue('X' . $initialRow, isset($domicilio['Provincia']) ? $domicilio['Provincia'] : '');
                    } else {
                        // Si la información del domicilio está como campos separados
                        $sheet->setCellValue('T' . $initialRow, isset($apoderado1['Domicilio - Tipo de Vía']) ? $apoderado1['Domicilio - Tipo de Vía'] : '');
                        $sheet->setCellValue('U' . $initialRow, isset($apoderado1['Domicilio - Nombre']) ? $apoderado1['Domicilio - Nombre'] : '');
                        $sheet->setCellValue('V' . $initialRow, isset($apoderado1['Domicilio - Número']) ? $apoderado1['Domicilio - Número'] : '');
                        $sheet->setCellValue('W' . $initialRow, isset($apoderado1['Domicilio - Localidad']) ? $apoderado1['Domicilio - Localidad'] : '');
                        $sheet->setCellValue('X' . $initialRow, isset($apoderado1['Domicilio - Provincia']) ? $apoderado1['Domicilio - Provincia'] : '');
                    }
                }

                // -----------------------------------------------------------Asignar datos del segundo apoderado
                if (isset($apoderados[1])) {
                    $apoderado2 = $apoderados[1];
                    $sheet->setCellValue('Z' . $initialRow, isset($apoderado2['Nombres']) ? $apoderado2['Nombres'] : '');
                    $sheet->setCellValue('AA' . $initialRow, isset($apoderado2['Apellidos']) ? $apoderado2['Apellidos'] : '');
                    $sheet->setCellValue('AB' . $initialRow, isset($apoderado2['Número DNI']) ? $apoderado2['Número DNI'] : '');
                    $sheet->setCellValue('AH' . $initialRow, isset($apoderado2['Tipo de apoderamiento']) ? $apoderado2['Tipo de apoderamiento'] : '');
                    // Verificar la estructura del domicilio del segundo apoderado
                    if (isset($apoderado2['Domicilio'])) {
                        // Si la información del domicilio está en un objeto 'Domicilio'
                        $domicilio = $apoderado2['Domicilio'];
                        $sheet->setCellValue('AC' . $initialRow, isset($domicilio['Tipo de Vía']) ? $domicilio['Tipo de Vía'] : '');
                        $sheet->setCellValue('AD' . $initialRow, isset($domicilio['Nombre']) ? $domicilio['Nombre'] : '');
                        $sheet->setCellValue('AE' . $initialRow, isset($domicilio['Número']) ? $domicilio['Número'] : '');
                        $sheet->setCellValue('AF' . $initialRow, isset($domicilio['Localidad']) ? $domicilio['Localidad'] : '');
                        $sheet->setCellValue('AG' . $initialRow, isset($domicilio['Provincia']) ? $domicilio['Provincia'] : '');
                    } else {
                        // Si la información del domicilio está como campos separados
                        $sheet->setCellValue('AC' . $initialRow, isset($apoderado2['Domicilio - Tipo de Vía']) ? $apoderado2['Domicilio - Tipo de Vía'] : '');
                        $sheet->setCellValue('AD' . $initialRow, isset($apoderado2['Domicilio - Nombre']) ? $apoderado2['Domicilio - Nombre'] : '');
                        $sheet->setCellValue('AE' . $initialRow, isset($apoderado2['Domicilio - Número']) ? $apoderado2['Domicilio - Número'] : '');
                        $sheet->setCellValue('AF' . $initialRow, isset($apoderado2['Domicilio - Localidad']) ? $apoderado2['Domicilio - Localidad'] : '');
                        $sheet->setCellValue('AG' . $initialRow, isset($apoderado2['Domicilio - Provincia']) ? $apoderado2['Domicilio - Provincia'] : '');
                    }
                }
                //----------------------------------------------------------- Asignar datos del tercer apoderado
if (isset($apoderados[2])) {
    $apoderado3 = $apoderados[2];
    $sheet->setCellValue('AI' . $initialRow, isset($apoderado3['Nombres']) ? $apoderado3['Nombres'] : '');
    $sheet->setCellValue('AJ' . $initialRow, isset($apoderado3['Apellidos']) ? $apoderado3['Apellidos'] : '');
    $sheet->setCellValue('AK' . $initialRow, isset($apoderado3['Número DNI']) ? $apoderado3['Número DNI'] : '');
    $sheet->setCellValue('AQ' . $initialRow, isset($apoderado3['Tipo de apoderamiento']) ? $apoderado3['Tipo de apoderamiento'] : '');

    // Verificar la estructura del domicilio del tercer apoderado
    if (isset($apoderado3['Domicilio'])) {
        // Si la información del domicilio está en un objeto 'Domicilio'
        $domicilio = $apoderado3['Domicilio'];
        $sheet->setCellValue('AL' . $initialRow, isset($domicilio['Tipo de Vía']) ? $domicilio['Tipo de Vía'] : '');
        $sheet->setCellValue('AM' . $initialRow, isset($domicilio['Nombre']) ? $domicilio['Nombre'] : '');
        $sheet->setCellValue('AN' . $initialRow, isset($domicilio['Número']) ? $domicilio['Número'] : '');
        $sheet->setCellValue('AO' . $initialRow, isset($domicilio['Localidad']) ? $domicilio['Localidad'] : '');
        $sheet->setCellValue('AP' . $initialRow, isset($domicilio['Provincia']) ? $domicilio['Provincia'] : '');
    } else {
        // Si la información del domicilio está como campos separados
        $sheet->setCellValue('AL' . $initialRow, isset($apoderado3['Domicilio - Tipo de Vía']) ? $apoderado3['Domicilio - Tipo de Vía'] : '');
        $sheet->setCellValue('AM' . $initialRow, isset($apoderado3['Domicilio - Nombre']) ? $apoderado3['Domicilio - Nombre'] : '');
        $sheet->setCellValue('AN' . $initialRow, isset($apoderado3['Domicilio - Número']) ? $apoderado3['Domicilio - Número'] : '');
        $sheet->setCellValue('AO' . $initialRow, isset($apoderado3['Domicilio - Localidad']) ? $apoderado3['Domicilio - Localidad'] : '');
        $sheet->setCellValue('AP' . $initialRow, isset($apoderado3['Domicilio - Provincia']) ? $apoderado3['Domicilio - Provincia'] : '');
    }
}

//---------------------------------------------------------------------- Asignar datos del cuarto apoderado
if (isset($apoderados[3])) {
    $apoderado4 = $apoderados[3];
    $sheet->setCellValue('AR' . $initialRow, isset($apoderado4['Nombres']) ? $apoderado4['Nombres'] : '');
    $sheet->setCellValue('AS' . $initialRow, isset($apoderado4['Apellidos']) ? $apoderado4['Apellidos'] : '');
    $sheet->setCellValue('AT' . $initialRow, isset($apoderado4['Número DNI']) ? $apoderado4['Número DNI'] : '');
    $sheet->setCellValue('AZ' . $initialRow, isset($apoderado4['Tipo de apoderamiento']) ? $apoderado4['Tipo de apoderamiento'] : '');

    // Verificar la estructura del domicilio del cuarto apoderado
    if (isset($apoderado4['Domicilio'])) {
        // Si la información del domicilio está en un objeto 'Domicilio'
        $domicilio = $apoderado4['Domicilio'];
        $sheet->setCellValue('AU' . $initialRow, isset($domicilio['Tipo de Vía']) ? $domicilio['Tipo de Vía'] : '');
        $sheet->setCellValue('AV' . $initialRow, isset($domicilio['Nombre']) ? $apoderado4['Domicilio']['Nombre'] : '');
        $sheet->setCellValue('AW' . $initialRow, isset($domicilio['Número']) ? $apoderado4['Domicilio']['Número'] : '');
        $sheet->setCellValue('AX' . $initialRow, isset($domicilio['Localidad']) ? $apoderado4['Domicilio']['Localidad'] : '');
        $sheet->setCellValue('AY' . $initialRow, isset($domicilio['Provincia']) ? $apoderado4['Domicilio']['Provincia'] : '');
    } else {
        // Si la información del domicilio está como campos separados
        $sheet->setCellValue('AU' . $initialRow, isset($apoderado4['Domicilio - Tipo de Vía']) ? $apoderado4['Domicilio - Tipo de Vía'] : '');
        $sheet->setCellValue('AV' . $initialRow, isset($apoderado4['Domicilio - Nombre']) ? $apoderado4['Domicilio - Nombre'] : '');
        $sheet->setCellValue('AW' . $initialRow, isset($apoderado4['Domicilio - Número']) ? $apoderado4['Domicilio - Número'] : '');
        $sheet->setCellValue('AX' . $initialRow, isset($apoderado4['Domicilio - Localidad']) ? $apoderado4['Domicilio - Localidad'] : '');
        $sheet->setCellValue('AY' . $initialRow, isset($apoderado4['Domicilio - Provincia']) ? $apoderado4['Domicilio - Provincia'] : '');
    }
}

                //------------------------------------------------- Accede al campo "Inscripción Registro Mercantil" dentro de $data -------------------------------------------//
                if (isset($data['Inscripción Reg. Mercantil'][0])) {
                    $registroMercantilData = $data['Inscripción Reg. Mercantil'][0];

                    // Mapeo para los campos de "Inscripción Registro Mercantil"
                    $registroMercantilMapping = [
                        'H' => 'Inscrito',
                        'I' => 'Hoja',
                        'J' => 'Tomo',
                        'K' => 'Inscripción',
                    ];

                    // Recorrer el mapeo y asignar los valores de "Inscripción Registro Mercantil" al Excel
                    foreach ($registroMercantilMapping as $column => $field) {
                        $cell = $column . $initialRow;
                        $value = isset($registroMercantilData[$field]) ? $registroMercantilData[$field] : '';

                        // Puedes agregar verificaciones adicionales o mensajes personalizados aquí si lo deseas
                        // Por ejemplo, si algún campo está vacío, puedes asignar un mensaje diferente.
                        // if ($field === 'Hoja' && empty($value)) {
                        //     $value = 'Hoja no disponible';
                        // }

                        $sheet->setCellValue($cell, is_string($value) ? $value : '');
                    }
                }
                //------------------------------------------------------- Accede al campo "Razón Social" dentro de $data--------------------------------------//
                if (isset($data['Razón Social'][0])) {
                    $razonSocialData = $data['Razón Social'][0];

                    // Mapeo para los campos de "Razón Social"
                    $razonSocialMapping = [
                        'E' => 'Denominación',
                        'F' => 'Domicilio social',
                    ];

                    // Recorrer el mapeo y asignar los valores de "Razón Social" al Excel
                    foreach ($razonSocialMapping as $column => $field) {
                        $cell = $column . $initialRow;
                        $value = isset($razonSocialData[$field]) ? $razonSocialData[$field] : '';

                        // Puedes agregar una verificación adicional aquí si lo deseas
                        // Por ejemplo, si el campo 'Domicilio social' no está presente, puedes asignar un mensaje diferente.
                        // if ($field === 'Domicilio social' && empty($value)) {
                        //     $value = 'Domicilio social no disponible';
                        // }

                        $sheet->setCellValue($cell, is_string($value) ? $value : '');
                    }
                }

                //------------------------------------ Accede al campo "Notario" dentro de $data-----------------------------------------------------//
                if (isset($data['Notario'])) {
                    $notarioData = $data['Notario'];

                    if (is_array($notarioData) && isset($notarioData[0])) {
                        // Si es un array de notarios con al menos un elemento, toma el primer elemento
                        $notario = $notarioData[0];

                        // Mapeo para los campos del notario
                        $notarioMapping = [
                            'L' => 'Nombre/Apellido',
                            'M' => 'Num. protocolo',
                            'N' => 'Fecha escritura',
                            'O' => 'Localidad',
                            'P' => 'Col. notarios',
                        ];

                        // Recorrer el mapeo y asignar los valores del notario al Excel
                        foreach ($notarioMapping as $column => $field) {
                            $cell = $column . $initialRow;
                            $value = isset($notario[$field]) ? $notario[$field] : '';
                            $sheet->setCellValue($cell, is_string($value) ? $value : '');
                        }
                    }
                }

                // Incrementa la variable de control de la fila
                $initialRow++;
            }
        }
    }
}
