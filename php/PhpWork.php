<?php

namespace php;

require_once __DIR__ . '../../vendor/autoload.php';

use Exception;
use Smalot\PdfParser\Parser;

class PhpWork
{
    private string $file;
    private ?Parser $pdf_parser = null;

    public function __construct(string $file)
    {
        $this->file = $file;
    }

    /**
     * @throws Exception
     */
    public function takeImages()
    {
        $this->pdf_parser = new Parser();

        $pdf = $this->pdf_parser->parseFile($this->file);
        $pages = $pdf->getPages();

        foreach ($pages as $page) {
            var_dump($page->getTextArray());
        }
    }

    public function __destruct()
    {
        if ($this->pdf_parser instanceof Parser) unset($this->pdf_parser);
    }
}

$file = 'F:/strateg/Лайно/інструкції/30324-ua.pdf';
$show_text_from_pdf = new PhpWork($file);
try {
    $show_text_from_pdf->takeImages();
} catch (Exception $e) {
    echo $e->getMessage();
}