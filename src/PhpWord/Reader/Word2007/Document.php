<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2018 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Reader\Word2007;

use PhpOffice\Common\XMLReader;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Html as HTMLParser;
/**
 * Document reader
 *
 * @since 0.10.0
 * @SuppressWarnings(PHPMD.UnusedPrivateMethod) For readWPNode
 */
class Document extends AbstractPart
{
    /**
     * PhpWord object
     *
     * @var \PhpOffice\PhpWord\PhpWord
     */
    private $phpWord;

    /**
     * Read document.xml.
     *
     * @param \PhpOffice\PhpWord\PhpWord $phpWord
     */
     
    private function getNumberingType($val, $zipFile) {
        $xmlFile = 'word/numbering.xml';
        $zip = new \ZipArchive();
        $zip->open($zipFile);
        $content = $zip->getFromName($xmlFile);
        $zip->close();

        if($val == '&gt;'){
            $val = 0;
        }
 
        $content = str_replace(array("\n", "\r"), array("", ""), $content);
        preg_match_all('/\<w\:num w\:numId\=\"'.$val.'\"\>\<w\:abstractNumId w\:val\=\"(.*?)\"\/\>\<\/w\:num\>/si', $content, $res);


        /*/"<w:num w:numId="2">
        <w:abstractNumId w:val="2"/>
        <w:lvlOverride w:ilvl="0">
        <w:lvl w:ilvl="0">
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/></w:lvl></w:lvlOverride></w:num><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>"
        /*/

        if(!isset($res[1][0])){
            return "b";
        }

        $abstract_num = $res[1][0];


        if($abstract_num == '&gt;'){
            $t="t";
            $abstract_num = 0;
        }
        
        if(is_numeric($abstract_num)){

            preg_match('/\<w\:abstractNum w\:abstractNumId\=\"'.$abstract_num.'\" w15\:restartNumberingAfterBreak\=\"0\">(.*?)\<\/w\:abstractNum\>/si', $content, $res2);
            
            if(!empty($res2[0])) {
                if(strpos($res2[0], 'decimal') !== FALSE) {
                    return "n";
                } else {
                    return "b";
                }            
            } else {
                preg_match('/\<w\:abstractNum w\:abstractNumId\=\"'.$abstract_num.'\">(.*?)\<\/w\:abstractNum\>/si', $content, $res2);
                if(strpos($res2[0], 'decimal') !== FALSE) {
                    return "n";
                } else {
                    return "b";
                }
            }

        }
        else{
            return "b";
        }

           

       

          
    } 
     
    public function read (PhpWord $phpWord)
    {
        $this->phpWord = $phpWord;
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);
        
        
        
        $readMethods = array('w:p' => 'readWPNode', 'w:tbl' => 'readTable', 'w:sectPr' => 'readWSectPrNode', 'w:sdt' => 'readSDTNode');
        $nodes = $xmlReader->getElements('w:body/*');
        
        if ($nodes->length > 0) {
            $section = $this->phpWord->addSection();
            $isOL = false;
            $isUL = false;
            foreach ($nodes as $key => $node) {
                $style = $xmlReader->getAttribute('w:val', $node, 'w:pPr/w:pStyle'); 
                $sectPrNodeArray = $xmlReader->getElements('w:r/w:t', $node);
                
                $isNumPR = $xmlReader->getElements('w:pPr/w:numPr', $node);
                //print_r($isNumPR);
                if(($style == 'Prrafodelista' || $isNumPR->length > 0)) {
                    $liType = intval($xmlReader->getAttribute('w:val', $node, 'w:pPr/w:numPr/w:numId'));
                    
                    if($key == 1112){
                        $test = "t";
                    }

                    $isNumbering = $this->getNumberingType($liType, $this->docFile);
                   
                    //print_r($liType);
                    if($sectPrNodeArray != null && $sectPrNodeArray->length > 0) {
                        if($isOL == false && $isNumbering == "n") {
                            $section->addText("[ol]");
                            $isOL = true;
                        }
                        if($isUL == false && $isNumbering == "b") {
                            $section->addText("[ul]");
                            $isUL = true;
                        }                        
                        $this->readOLNode($xmlReader, $node, $section);                         
                    }     
                } else if (isset($readMethods[$node->nodeName])) {
                    if($isOL == true) {
                        $section->addText("[/ol]");
                        $isOL = false;
                    }
                    if($isUL == true) {
                        $section->addText("[/ul]");
                        $isUL = false;
                    }
                    $readMethod = $readMethods[$node->nodeName];
                    $this->$readMethod($xmlReader, $node, $section);
                }
            }
            if($isOL == true) {
                $section->addText("[/ol]");
                $isOL = false;
            }            
            //\PhpOffice\PhpWord\Shared\Html::addHtml($section, '<ol><li>HOLA</li></ol>', true, false);
        }
        
    }

    public function readold(PhpWord $phpWord)
    {
        $this->phpWord = $phpWord;
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);
        $readMethods = array('w:p' => 'readWPNode', 'w:tbl' => 'readTable', 'w:sectPr' => 'readWSectPrNode');

        $nodes = $xmlReader->getElements('w:body/*');
        if ($nodes->length > 0) {
            $section = $this->phpWord->addSection();
            foreach ($nodes as $node) {
                if (isset($readMethods[$node->nodeName])) {
                    $readMethod = $readMethods[$node->nodeName];
                    $this->$readMethod($xmlReader, $node, $section);
                }
            }
        }
    }

    /**
     * Read header footer.
     *
     * @param array $settings
     * @param \PhpOffice\PhpWord\Element\Section &$section
     */
    private function readHeaderFooter($settings, Section &$section)
    {
        $readMethods = array('w:p' => 'readParagraph', 'w:tbl' => 'readTable');

        if (is_array($settings) && isset($settings['hf'])) {
            foreach ($settings['hf'] as $rId => $hfSetting) {
                if (isset($this->rels['document'][$rId])) {
                    list($hfType, $xmlFile, $docPart) = array_values($this->rels['document'][$rId]);
                    $addMethod = "add{$hfType}";
                    $hfObject = $section->$addMethod($hfSetting['type']);

                    // Read header/footer content
                    $xmlReader = new XMLReader();
                    $xmlReader->getDomFromZip($this->docFile, $xmlFile);
                    $nodes = $xmlReader->getElements('*');
                    if ($nodes->length > 0) {
                        foreach ($nodes as $node) {
                            if (isset($readMethods[$node->nodeName])) {
                                $readMethod = $readMethods[$node->nodeName];
                                $this->$readMethod($xmlReader, $node, $hfObject, $docPart);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Read w:sectPr
     *
     * @param \PhpOffice\Common\XMLReader $xmlReader
     * @param \DOMElement $domNode
     * @ignoreScrutinizerPatch
     * @return array
     */
    private function readSectionStyle(XMLReader $xmlReader, \DOMElement $domNode)
    {
        $styleDefs = array(
            'breakType'     => array(self::READ_VALUE, 'w:type'),
            'pageSizeW'     => array(self::READ_VALUE, 'w:pgSz', 'w:w'),
            'pageSizeH'     => array(self::READ_VALUE, 'w:pgSz', 'w:h'),
            'orientation'   => array(self::READ_VALUE, 'w:pgSz', 'w:orient'),
            'colsNum'       => array(self::READ_VALUE, 'w:cols', 'w:num'),
            'colsSpace'     => array(self::READ_VALUE, 'w:cols', 'w:space'),
            'marginTop'     => array(self::READ_VALUE, 'w:pgMar', 'w:top'),
            'marginLeft'    => array(self::READ_VALUE, 'w:pgMar', 'w:left'),
            'marginBottom'  => array(self::READ_VALUE, 'w:pgMar', 'w:bottom'),
            'marginRight'   => array(self::READ_VALUE, 'w:pgMar', 'w:right'),
            'headerHeight'  => array(self::READ_VALUE, 'w:pgMar', 'w:header'),
            'footerHeight'  => array(self::READ_VALUE, 'w:pgMar', 'w:footer'),
            'gutter'        => array(self::READ_VALUE, 'w:pgMar', 'w:gutter')
        );

        $styles = $this->readStyleDefs($xmlReader, $domNode, $styleDefs);
        // Header and footer
        // @todo Cleanup this part
        $nodes = $xmlReader->getElements('*', $domNode);
        foreach ($nodes as $node) {
            if ($node->nodeName == 'w:headerReference' || $node->nodeName == 'w:footerReference') {
                $id = $xmlReader->getAttribute('r:id', $node);
                $styles['hf'][$id] = array(
                    'method' => str_replace('w:', '', str_replace('Reference', '', $node->nodeName)),
                    'type'   => $xmlReader->getAttribute('w:type', $node),
                );
            }
        }

        return $styles;
    }

    /**
     * Read w:p node.
     *
     * @param \PhpOffice\Common\XMLReader $xmlReader
     * @param \DOMElement $node
     * @param \PhpOffice\PhpWord\Element\Section &$section
     *
     * @todo <w:lastRenderedPageBreak>
     */
    private function readWPNode(XMLReader $xmlReader, \DOMElement $node, Section &$section)
    {
        // Page break
        if ($xmlReader->getAttribute('w:type', $node, 'w:r/w:br') == 'page') {
            $section->addPageBreak(); // PageBreak
        }

        // Paragraph
        $this->readParagraph($xmlReader, $node, $section);

        //$node->getFontStyle()->setLineHeight(3.0);
        $isSingle = false;
        //
        
        $sectPrNodeUnderline = $xmlReader->getElement('w:r/w:rPr/w:u', $node);
        if($sectPrNodeUnderline != null) {
            foreach($sectPrNodeUnderline->attributes as $attr) {
                if($attr->nodeName == "w:val" && $attr->nodeValue == "single") {
                    $isSingle = true;
                }
            }
        }
        
        $style = $this->readSectionStyle($xmlReader, $node);
        if($isSingle) {
            $style['underline'] = 'single';
        }
        
        //$style['lineHeight'] = 1.5;
        //$lineHeight;
        //echo 'LH: ' . $lineHeight."\n";
        //print_r($section->getStyle());
        $section->setStyle($style);
        // Section properties
        if ($xmlReader->elementExists('w:pPr/w:sectPr', $node)) {
            $sectPrNode = $xmlReader->getElement('w:pPr/w:sectPr', $node);
            if ($sectPrNode !== null) {
                $this->readWSectPrNode($xmlReader, $sectPrNode, $section);
            }
            $section = $this->phpWord->addSection();
        }
    }
    
    private function readOLNode(XMLReader $xmlReader, \DOMElement $node, Section &$section) {
        $this->readParagraph($xmlReader, $node, $section);
        
        // Section properties
        if ($xmlReader->elementExists('w:r', $node)) {
            $section = $this->phpWord->addSection();
            //$table = $section->addTable('myOwnTableStyle');
            $sectPrNodeArray = $xmlReader->getElements('w:r/w:t', $node);
            $itemInArray = array();
            foreach($sectPrNodeArray as $itemNode) {
                //$section->addListItem($sectPrNode->nodeValue);
                $itemInArray[] = $itemNode->nodeValue;
            }
            $section->addText("[li]" . implode('', $itemInArray) . "[/li]");
            //$this->readWSectPrNode($xmlReader, $sectPrNode, $section);
        }
        
        
    }
    
    private function readSDTNode(XMLReader $xmlReader, \DOMElement $node, Section &$section) {
        $this->readParagraph($xmlReader, $node, $section);

        // Section properties
        if ($xmlReader->elementExists('w:sdtContent', $node)) {
            $section = $this->phpWord->addSection();
            //$table = $section->addTable('myOwnTableStyle');
            $sectPrNode = $xmlReader->getElement('w:sdtContent', $node);
            foreach($sectPrNode->childNodes as $key => $nodeTR) {
                ///$table->addRow();
                //$table->addCell(2500, null)->addText($nodeTR->nodeValue);
                /*if($nodeTR->childNodes->length > 0) {
                    foreach($nodeTR->childNodes as $child) {
                        print_r($child);
                        print_r($child->attributes);
                    }
                }*/
                //print_r($nodeTR);
                //print_r($nodeTR->attributes);
                $section->addText($nodeTR->nodeValue);
            }
            $this->readWSectPrNode($xmlReader, $sectPrNode, $section);
        }        
    }

    /**
     * Read w:sectPr node.
     *
     * @param \PhpOffice\Common\XMLReader $xmlReader
     * @param \DOMElement $node
     * @param \PhpOffice\PhpWord\Element\Section &$section
     */
    private function readWSectPrNode(XMLReader $xmlReader, \DOMElement $node, Section &$section)
    {
        $style = $this->readSectionStyle($xmlReader, $node);
        $section->setStyle($style);
        $this->readHeaderFooter($style, $section);
    }

    public function getphpWord(){
        return $this->phpWord;
    }

}
