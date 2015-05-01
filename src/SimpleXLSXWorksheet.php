<?php
/**
 * Created by IntelliJ IDEA.
 * User: eugeny.mikhalev
 * Date: 07.07.14
 * Time: 15:02
 * To change this template use File | Settings | File Templates.
 */
namespace emikhalev/SimpleXLSX;

class SimpleXLSXWorksheet
{
    private $name = null;
    private $cells = [];
    private $colsWidth = [];
    private $rowsHeight = [];
    private $merges = [];

    const WIDTH_DEFAULT = 10;
    const HEIGHT_DEFAULT = 3;

    public function getName()
    {
        return $this->name;
    }

    public function setName($name)
    {
        $this->name = $name;
    }

    public function __construct($name = null)
    {
        $this->setName( $name );
    }

    public function getCells()
    {
        return $this->cells;
    }

    public function setCellVal($cellRow, $cellCol, $val)
    {
        $this->cells[$cellRow][$cellCol]['value'] = $val;
        if ( !isset($this->cells[$cellRow][$cellCol]['style']) )
        {
            $this->cells[$cellRow][$cellCol]['style']='__default';
        }
    }

    public function getCellVal($cellRow, $cellCol)
    {
        return isset($this->cells[$cellRow][$cellCol]['value']) ? $this->cells[$cellRow][$cellCol]['value'] : null;
    }

    public function setCellStyle($cellRow, $cellCol, $style)
    {
        $this->cells[$cellRow][$cellCol]['style']=$style;
        if ( !isset($this->cells[$cellRow][$cellCol]['value']) )
        {
            $this->cells[$cellRow][$cellCol]['value'] = '';
        }
    }

    public function getCellStyle($cellRow, $cellCol)
    {
        return isset($this->cells[$cellRow][$cellCol]['style']) ? $this->cells[$cellRow][$cellCol]['style'] : null;
    }

    public function setColWidth($colNum, $width = self::WIDTH_DEFAULT)
    {
        $this->colsWidth[$colNum] = $width;
    }

    public function getColWidth($colNum)
    {
        return $this->colsWidth[$colNum];
    }

    public function getColsWidth()
    {
        return $this->colsWidth;
    }

    public function setRowHeight($rowNum, $rowHeight = self::HEIGHT_DEFAULT)
    {
        $this->rowsHeight[$rowNum] = $rowHeight;
    }

    public function getRowHeight($rowNum)
    {
        return $this->rowsHeight[$rowNum];
    }

    public function mergeCells($startRow, $startCol, $endRow, $endCol)
    {
        $this->merges[] = [
            'startCol' => $startCol,
            'endCol' => $endCol,
            'startRow' => $startRow,
            'endRow' => $endRow,
        ];
    }

    public function getMerges()
    {
        return $this->merges;
    }
}