<?php

namespace Threshold\PhpExcel\Writer\Helper;

use SimpleXMLElement;

class ExtendedSimpleXMLElement extends SimpleXMLElement
{
    public function prependChild($name, $value = "")
    {
        $dom = dom_import_simplexml($this);
        $new = $dom->insertBefore(
            $dom->ownerDocument->createElement($name, $value),
            $dom->firstChild
        );
        
        return simplexml_import_dom($new, get_class($this));
    }
    
    public function insertChild($name, $value = "", int $position = 0)
    {
        $dom = dom_import_simplexml($this);
        $new = $dom->insertBefore(
            $dom->ownerDocument->createElement($name, $value),
            $dom->childNodes[$position]
        );
        
        return simplexml_import_dom($new, get_class($this));
    }
}