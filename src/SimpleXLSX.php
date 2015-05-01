<?php
/**
 * Created by IntelliJ IDEA.
 * User: eugeny.mikhalev
 * Date: 07.07.14
 * Time: 11:43
 * To change this template use File | Settings | File Templates.
 */
namespace emikhalev\SimpleXLSX;

class SimpleXLSX
{

    public function __construct()
    {
    }

    public function createWorkbook()
    {
        $workbook = new SimpleXLSXWorkbook();
        return $workbook;
    }
}
