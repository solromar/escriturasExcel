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
        $jsonFolderPath = $this->getParameter('kernel.project_dir') . '/public/1-prueba';

        // Obtener la lista de archivos JSON en la carpeta
        $jsonFiles = glob($jsonFolderPath . '/*.json');


        // Ruta del archivo Excel predefinido
        $excelFilePath = $this->getParameter('kernel.project_dir') . '/public/archivosExcel/otra.xlsx';

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
        // Obtener el valor de fileName de la sección "Categorizador_GPT4 8K Azure"
        // $fileName = $data['processes']['Categorizador_GPT4 8K Azure']['result']['fileName'];

        // Agregar el valor de fileName a la columna A
        //$sheet->setCellValue('A' . $initialRow, $fileName);


        $answer = null;

        if (isset($data['processes']['GPT4 8K Azure']['result']['answer'])) {
            $answer = $data['processes']['GPT4 8K Azure']['result']['answer'];
        } else {
            foreach ($data as $key => $value) {
                if (isset($value['result']) && is_array($value['result'])) {
                    foreach ($value['result'] as $result) {
                        if (isset($result['model']) && $result['model'] === 'GPT4 8K Azure (gpt-4)' && isset($result['answer'])) {
                            $answer = $result['answer'];
                            break 2; // Salir de ambos bucles si se encuentra el resultado deseado
                        }
                    }
                }
            }
        }

        // Ahora puedes acceder a las variables dentro de $answer
        if ($answer !== null) {


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
            if (isset($answer['Razón Social'])) {
                $razonSocialData = $answer['Razón Social'];

                if (is_array($razonSocialData) && isset($razonSocialData[0])) {
                    // Si es un array de notarios con al menos un elemento, toma el primer elemento
                    $razonSocial = $razonSocialData[0];
                } else {
                    // Si es un objeto de notario o un array con un solo elemento, usa directamente ese objeto
                    $razonSocial = (array)($razonSocialData[0] ?? $razonSocialData);
                }

                // Ahora puedes acceder a los campos de notario y asignarlos a las celdas de Excel
                $sheet->setCellValue('E' . $initialRow, $razonSocial['Denominación']);
                // Verificar si 'Domicilio social' está presente antes de asignarlo
                if (isset($razonSocial['Domicilio social'])) {
                    $sheet->setCellValue('F' . $initialRow, $razonSocial['Domicilio social']);
                } else {
                    // Asignar un mensaje cuando 'Domicilio social' no está presente
                    $sheet->setCellValue('F' . $initialRow, 'Domicilio social no disponible');
                    
                }
            }

            //-----------------------------------Inscripcion Registro Mercantil-------------------------------//
            if (isset($answer['Inscripción Reg. Mercantil'])) {
                $registroData = $answer['Inscripción Reg. Mercantil'];

                if (is_array($registroData) && isset($registroData[0])) {
                    // Si es un array de notarios con al menos un elemento, toma el primer elemento
                    $registroMercantil = $registroData[0];
                } else {
                    // Si es un objeto de notario o un array con un solo elemento, usa directamente ese objeto
                    $registroMercantil = (array)($registroData[0] ?? $registroData);
                }

                // Ahora puedes acceder a los campos de notario y asignarlos a las celdas de Excel
                $sheet->setCellValue('H' . $initialRow, $registroMercantil['Inscrito']);
                $sheet->setCellValue('I' . $initialRow, $registroMercantil['Hoja']);
                $sheet->setCellValue('J' . $initialRow, $registroMercantil['Tomo']);
                $sheet->setCellValue('K' . $initialRow, $registroMercantil['Inscripción']);
            }
            
            //---------------------------------------Notario----------------------------------------//
            if (isset($answer['Notario'])) {
                $notarioData = $answer['Notario'];

                if (is_array($notarioData) && isset($notarioData[0])) {
                    // Si es un array de notarios con al menos un elemento, toma el primer elemento
                    $notario = $notarioData[0];
                } else {
                    // Si es un objeto de notario o un array con un solo elemento, usa directamente ese objeto
                    $notario = (array)($notarioData[0] ?? $notarioData);
                }

                // Ahora puedes acceder a los campos de notario y asignarlos a las celdas de Excel
                $sheet->setCellValue('L' . $initialRow, $notario['Nombre/Apellido']);
                $sheet->setCellValue('M' . $initialRow, $notario['Num. protocolo']);
                $sheet->setCellValue('N' . $initialRow, $notario['Fecha escritura']);
                $sheet->setCellValue('O' . $initialRow, $notario['Localidad']);
                $sheet->setCellValue('P' . $initialRow, $notario['Col. notarios']);
            }
            //--------------------------------------Apoderados hasta 8---------------------------------//
            if (isset($answer['Apoderado'][0])) {
                $apoderado1 = $answer['Apoderado'][0];
                $sheet->setCellValue('Q' . $initialRow, $apoderado1['Nombres']);
                $sheet->setCellValue('R' . $initialRow, $apoderado1['Apellidos']);
                $sheet->setCellValue('S' . $initialRow, $apoderado1['Número DNI']);
                // Verificar la estructura del domicilio
                if (isset($apoderado1['Domicilio'])) {
                    // Si la información del domicilio está en un objeto 'Domicilio'
                    $domicilio = $apoderado1['Domicilio'];
                    $sheet->setCellValue('T' . $initialRow, $domicilio['Tipo de Vía']);
                    $sheet->setCellValue('U' . $initialRow, $domicilio['Nombre']);
                    $sheet->setCellValue('V' . $initialRow, $domicilio['Número']);
                    $sheet->setCellValue('W' . $initialRow, $domicilio['Localidad']);
                    $sheet->setCellValue('X' . $initialRow, $domicilio['Provincia']);
                } else {
                    // Si la información del domicilio está como campos separados
                    $sheet->setCellValue('T' . $initialRow, $apoderado1['Domicilio - Tipo de Vía']);
                    $sheet->setCellValue('U' . $initialRow, $apoderado1['Domicilio - Nombre']);
                    $sheet->setCellValue('V' . $initialRow, $apoderado1['Domicilio - Número']);
                    $sheet->setCellValue('W' . $initialRow, $apoderado1['Domicilio - Localidad']);
                    $sheet->setCellValue('X' . $initialRow, $apoderado1['Domicilio - Provincia']);
                }
                $sheet->setCellValue('Y' . $initialRow, $apoderado1['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][1])) {
                $apoderado2 = $answer['Apoderado'][1];
                $sheet->setCellValue('Z' . $initialRow, $apoderado2['Nombres']);
                $sheet->setCellValue('AA' . $initialRow, $apoderado2['Apellidos']);
                $sheet->setCellValue('AB' . $initialRow, $apoderado2['Número DNI']);
                // Verificar la estructura del domicilio
                if (isset($apoderado2['Domicilio'])) {
                    // Si la información del domicilio está en un objeto 'Domicilio'
                    $domicilio = $apoderado2['Domicilio'];
                    $sheet->setCellValue('AC' . $initialRow, $domicilio['Tipo de Vía']);
                    $sheet->setCellValue('AD' . $initialRow, $domicilio['Nombre']);
                    $sheet->setCellValue('AE' . $initialRow, $domicilio['Número']);
                    $sheet->setCellValue('AF' . $initialRow, $domicilio['Localidad']);
                    $sheet->setCellValue('AG' . $initialRow, $domicilio['Provincia']);
                } else {
                    // Si la información del domicilio está como campos separados
                    $sheet->setCellValue('AC' . $initialRow, $apoderado2['Domicilio - Tipo de Vía']);
                    $sheet->setCellValue('AD' . $initialRow, $apoderado2['Domicilio - Nombre']);
                    $sheet->setCellValue('AE' . $initialRow, $apoderado2['Domicilio - Número']);
                    $sheet->setCellValue('AF' . $initialRow, $apoderado2['Domicilio - Localidad']);
                    $sheet->setCellValue('AG' . $initialRow, $apoderado2['Domicilio - Provincia']);
                }
                $sheet->setCellValue('AH' . $initialRow, $apoderado2['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][2])) {
                $apoderado3 = $answer['Apoderado'][2];
                $sheet->setCellValue('AI' . $initialRow, $apoderado3['Nombres']);
                $sheet->setCellValue('AJ' . $initialRow, $apoderado3['Apellidos']);
                $sheet->setCellValue('AK' . $initialRow, $apoderado3['Número DNI']);
                // Verificar la estructura del domicilio
                if (isset($apoderado3['Domicilio'])) {
                    // Si la información del domicilio está en un objeto 'Domicilio'
                    $domicilio = $apoderado3['Domicilio'];
                    $sheet->setCellValue('AL' . $initialRow, $domicilio['Tipo de Vía']);
                    $sheet->setCellValue('AM' . $initialRow, $domicilio['Nombre']);
                    $sheet->setCellValue('AN' . $initialRow, $domicilio['Número']);
                    $sheet->setCellValue('AO' . $initialRow, $domicilio['Localidad']);
                    $sheet->setCellValue('AP' . $initialRow, $domicilio['Provincia']);
                } else {
                    // Si la información del domicilio está como campos separados
                    $sheet->setCellValue('AL' . $initialRow, $apoderado3['Domicilio - Tipo de Vía']);
                    $sheet->setCellValue('AM' . $initialRow, $apoderado3['Domicilio - Nombre']);
                    $sheet->setCellValue('AN' . $initialRow, $apoderado3['Domicilio - Número']);
                    $sheet->setCellValue('AO' . $initialRow, $apoderado3['Domicilio - Localidad']);
                    $sheet->setCellValue('AP' . $initialRow, $apoderado3['Domicilio - Provincia']);
                }
                $sheet->setCellValue('AQ' . $initialRow, $apoderado3['Tipo de apoderamiento']);
            }
            if (isset($answer['Apoderado'][3])) {
                $apoderado4 = $answer['Apoderado'][3];
                $sheet->setCellValue('AR' . $initialRow, $apoderado4['Nombres']);
                $sheet->setCellValue('AS' . $initialRow, $apoderado4['Apellidos']);
                $sheet->setCellValue('AT' . $initialRow, $apoderado4['Número DNI']);
                // Verificar la estructura del domicilio
                if (isset($apoderado4['Domicilio'])) {
                    // Si la información del domicilio está en un objeto 'Domicilio'
                    $domicilio = $apoderado4['Domicilio'];
                    $sheet->setCellValue('AU' . $initialRow, $domicilio['Tipo de Vía']);
                    $sheet->setCellValue('AV' . $initialRow, $domicilio['Nombre']);
                    $sheet->setCellValue('AW' . $initialRow, $domicilio['Número']);
                    $sheet->setCellValue('AX' . $initialRow, $domicilio['Localidad']);
                    $sheet->setCellValue('AY' . $initialRow, $domicilio['Provincia']);
                } else {
                    // Si la información del domicilio está como campos separados
                    $sheet->setCellValue('AU' . $initialRow, $apoderado4['Domicilio - Tipo de Vía']);
                    $sheet->setCellValue('AV' . $initialRow, $apoderado4['Domicilio - Nombre']);
                    $sheet->setCellValue('AW' . $initialRow, $apoderado4['Domicilio - Número']);
                    $sheet->setCellValue('AX' . $initialRow, $apoderado4['Domicilio - Localidad']);
                    $sheet->setCellValue('AY' . $initialRow, $apoderado4['Domicilio - Provincia']);
                }
                $sheet->setCellValue('AZ' . $initialRow, $apoderado4['Tipo de apoderamiento']);
            }
        } else {
            // No se encontró el resultado deseado en ninguna de las estructuras
        }
    }
}
