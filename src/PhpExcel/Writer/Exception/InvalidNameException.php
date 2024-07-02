<?php

namespace Threshold\PhpExcel\Writer\Exception;

use Exception;
use Threshold\PhpExcel\Writer\Entity\Style\BorderPart;

class InvalidNameException extends Exception
{
    public function __construct($name)
    {
        $msg = '%s is not a valid name identifier for a border. Valid identifiers are: %s.';

        parent::__construct(\sprintf($msg, $name, \implode(',', BorderPart::getAllowedNames())));
    }
}
