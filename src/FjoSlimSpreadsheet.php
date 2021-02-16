<?php

namespace fjourneau\Spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use Slim\Http\Response;

/**
 * Classe améliorée pour sites FJO pour génération extract XLS avec gestion Slim 3 Response.
 *
 * @package fjourneau\Spreadsheet
 * @author fJourneau
 */
class FjoSlimSpreadsheet extends FjoSpreadsheet
{

    /**
     * Génère le fichier XLSX à télécharger ou sauver sur le serveur
     *
     * @param  Response $response (Slim 3 Response)
     * @param  string $filename Nom du fichier à télécharger ou endroit où sauvegarder sur le serveur
     * @return Response
     */
    public function generateFileInResponse(Response $response, $filename = 'file.xlsx'): Response
    {
        $headers = $this->getHttpHeaders($filename);

        foreach ($headers as $key => $val) {
            $response = $response->withHeader($key, $val);
        }

        $writer = IOFactory::createWriter($this, $this->getFilenameInfos($filename)['writer']);

        ob_start();
        $writer->save('php://output');
        $file_content = ob_get_clean();

        return $response->write($file_content);
    }
}
