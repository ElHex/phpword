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

namespace PhpOffice\PhpWord\Writer\HTML\Element;

use PhpOffice\PhpWord\Settings;

/**
 * TextRun element HTML writer
 *
 * @since 0.10.0
 */
class Title extends AbstractElement
{
    /**
     * Write heading
     *
     * @return string
     */
    public function write()
    {
        if (!$this->element instanceof \PhpOffice\PhpWord\Element\Title) {
            return '';
        }

        $fontStyle = null;
        $style = "";


        $tag = 'h' . $this->element->getDepth();

        $text = $this->element->getText();
        if (is_string($text)) {

            if (Settings::isOutputEscapingEnabled()) {
                $text = $this->escaper->escapeHtml($text);
            }

            if(isset($this->element->fontStyle)){

                $fontStyle = $this->element->fontStyle->getStyleValues();
    
                if($fontStyle['basic']['size'] != null){
                    $style .= "font-size: ".$fontStyle['basic']['size']."pt;";
                }
    
            }

        } elseif ($text instanceof \PhpOffice\PhpWord\Element\AbstractContainer) {
            $writer = new Container($this->parentWriter, $text);
            $text = $writer->write();
        }

        if($style!=""){

            $content ="<{$tag}><span style='{$style}'>{$text}</span></{$tag}>" . PHP_EOL;
            //$content ="<font style='{$style}'>{$text}</font>" . PHP_EOL;

        }
        else{

            $content = "<{$tag}>{$text}</{$tag}>" . PHP_EOL;

        }        

        return $content;
    }
}
