<?php
/**
 * Created by IntelliJ IDEA.
 * User: eugeny.mikhalev
 * Date: 07.07.14
 * Time: 15:01
 * To change this template use File | Settings | File Templates.
 */
namespace emikhalev\SimpleXLSX;

class SimpleXLSXWorkbook
{
    const START_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

    // Files locations
    const WORKBOOK_PATH = 'xl/';
    const WORKBOOK_XML_PATH = 'xl/';
    const WORKSHEET_PATH = 'xl/worksheets/';
    const WORKSHEET_XML_PATH = 'worksheets/';
    const WORKBOOK_RELS_PATH = 'xl/_rels/';
    const PL_RELATIONSHIP_PATH = '_rels/';
    const STYLES_PATH = 'xl/';
    const STYLES_XML_PATH = '';

    // Content constants
    const START_CONTENT_TYPES = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    const END_CONTENT_TYPES = '</Types>';
    const DEF_EXTENSION_RELS = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
    const PART_NAME_WORKBOOK = '<Override PartName="{workbook_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />';
    const PART_NAME_WORKSHEET= '<Override PartName="{worksheet_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />';
    const PART_NAME_STYLES = '<Override PartName="{styles_path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />';

    // Package Level (PL) Relationship constants
    const RELS_START = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    const RELS_END = '</Relationships>';
    const REL_TO_WB = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{workbookPath}" />';

    // Workbook constants
    const WB_START = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    const WB_SHEETS_START = '<sheets>';
    const WB_SHEET = '<sheet name="{sheetName}" sheetId="{sheetId}" r:id="{rId}" />';
    const WB_SHEETS_END = '</sheets>';
    const WB_END = '</workbook>';

    // WB RELS constants
    const WB_RELS_TO_WS = '<Relationship Id="{rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{sheetPath}" />';
    const WB_RELS_TO_STYLES = '<Relationship Id="{rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="{stylePath}" />';

    // Styles constants
    const STYLES_START = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    // -- Styles fonts
    const STYLES_FONTS_START = '<fonts count="{count}">';
    const STYLES_FONT_START = '<font>';
    const STYLES_FONT_SIZE = '<sz val="{size}"/>';
    const STYLES_FONT_NAME = '<name val="{name}"/>';
    const STYLES_FONT_FAMILY = '<family val="{family}"/>';
    const STYLES_FONT_STYLE_BOLD = '<b/>';
    const STYLES_FONT_STYLE_ITALIC = '<i/>';
    const STYLES_FONT_COLOR = '<color indexed="{colorIndex}"/>';
    const STYLES_FONT_END = '</font>';
    const STYLES_FONTS_END = '</fonts>';
    // -- Styles colors
    const STYLES_COLORS_START = '<colors>';
    const STYLES_COLOR_INDEXED_START = '<indexedColors>';
    const STYLES_COLOR_RGB = '<rgbColor rgb="{rgbColor}"/>';
    const STYLES_COLOR_INDEXED_END = '</indexedColors>';
    const STYLES_COLORS_END = '</colors>';
    // -- Styles borders
    const STYLES_BORDERS_START = '<borders count="{count}">';
    const STYLES_BORDER_START = '<border>';
    const STYLES_BORDER_LEFT = '<left style="{style}"/>';
    const STYLES_BORDER_RIGHT = '<right style="{style}"/>';
    const STYLES_BORDER_TOP='<top style="{style}"/>';
    const STYLES_BORDER_BOTTOM='<bottom style="{style}"/>';
    const STYLES_BORDER_DIAGONAL='<diagonal style="{style}">';
    const STYLES_BORDER_END = '</border>';
    const STYLES_BORDERS_END = '</borders>';
    // -- Fills
    const STYLES_FILLS_START = '<fills count="{count}">';
    const STYLES_FILLS_EXCEL_DEFAULT = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>';
    const STYLES_FILL_START = '<fill>';
    const STYLES_FILL_PATTERN_START = '<patternFill patternType="solid">';
    const STYLES_FILL_FGCOLOR = '<fgColor rgb="{rgb}"/>';
    const STYLES_FILL_BGCOLOR = '<bgColor rgb="{rgb}"/>';
    const STYLES_FILL_PATTERN_END = '</patternFill>';
    const STYLES_FILL_END = '</fill>';
    const STYLES_FILLS_END = '</fills>';
    // -- Styles xfs
    const STYLES_XFS_CELLS_START = '<cellXfs count="{count}">';
    const STYLES_XFS_START = '<xf fontId="{fontId}" borderId="{borderId}" fillId="{fillId}" applyBorder="1" applyFont="1" applyAlignment="1" applyFill="1">';
    const STYLES_XFS_ALIGNMENT = '<alignment horizontal="{horizontal}" vertical="{vertical}" wrapText="{isWrapText}"/>';
    const STYLES_XFS_END = '</xf>';
    const STYLES_XFS_CELLS_END = '</cellXfs>';
    const STYLES_END = '</styleSheet>';

    // Worksheet constants
    const WS_START = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    const WS_PR_STUB = '<sheetPr />';
    const WS_DATA_STUB = '<sheetData />';
    const WS_DATA_START = '<sheetData>';
    const WS_ROW_START = '<row r="{rowNum}">';
    const WS_COLS_START = '<cols>';
    const WS_COL = '<col min="{colNum}" max="{colNum}" customWidth="1" width="{width}" />';
    const WS_COLS_END = '</cols>';
    const WS_CELL_START = '<c r="{cellName}" t="inlineStr" s="{styleNum}">';
    const WS_CELL_VALUE = '<is><t>{cellVal}</t></is>';
    const WS_CELL_END = '</c>';
    const WS_ROW_END = '</row>';
    const WS_DATA_END = '</sheetData>';
    const WS_END = '</worksheet>';
    const WS_MERGES_START = '<mergeCells>';
    const WS_MERGE = '<mergeCell ref="{startCell}:{endCell}"/>';
    const WS_MERGES_END = '</mergeCells>';

    private $worksheets = [];
    private $styles = [];

    public function __construct()
    {
        for  ($i=0; $i<10; $i++)
        {
            $this->createStyle("stubForExcel".$i);
        }
        $this->createStyle('__default');
    }

    public function createWorksheet()
    {
        $worksheet = new SimpleXLSXWorksheet();
        $this->worksheets[] = $worksheet;
        return $worksheet;
    }

    public function createStyle($name)
    {
        $styleId = count($this->styles);
        $this->styles[$name] = new SimpleXLSXStyle( $styleId );
        return $this->styles[$name];
    }

    public function getStyle($name)
    {
        return isset($this->styles[$name]) ? $this->styles[$name] : null;
    }

    public function save($filename)
    {

        // Creating xml parts
        $ctype = $this->makeContentTypeItem();
        $plrels = $this->makePackageLevelRelationship();
        $wb = $this->makeWorkBook();
        $wbRels = $this->makeWorkbookRelationship();
        $styles = $this->makeStyles();

        // Zipping
        $zip = new \ZipArchive();
        if ( $zip->open($filename, \ZIPARCHIVE::CREATE)!==TRUE )
        {
            return false;
        }

        $zip->addFromString('[Content_Types].xml', $ctype);
        $zip->addFromString(self::PL_RELATIONSHIP_PATH.'.rels', $plrels);
        $zip->addFromString(self::WORKBOOK_PATH.'workbook.xml', $wb);
        $zip->addFromString(self::WORKBOOK_RELS_PATH.'workbook.xml.rels', $wbRels);
        $zip->addFromString(self::STYLES_PATH.'styles.xml', $styles);

        // Making worksheets and zipping
        reset($this->worksheets);
        for ($i=1; $i<=count($this->worksheets); $i++)
        {
            $ws = $this->makeWorksheet( current($this->worksheets) );
            $zip->addFromString(self::WORKSHEET_PATH."sheet$i.xml", $ws);
            next($this->worksheets);
        }

        return $zip->close();
    }

    private function makeContentTypeItem()
    {
        $xml  = self::START_XML;
        $xml .= self::START_CONTENT_TYPES;
        $xml .= self::DEF_EXTENSION_RELS;
        $xml .= str_replace( '{workbook_path}','/'.self::WORKBOOK_PATH.'workbook.xml', self::PART_NAME_WORKBOOK );
        $xml .= str_replace( '{styles_path}','/'.self::STYLES_PATH.'styles.xml', self::PART_NAME_STYLES );

        for ($i=1; $i<=count($this->worksheets); $i++)
        {
            $xml .= str_replace( '{worksheet_path}','/'.self::WORKSHEET_PATH."sheet$i.xml", self::PART_NAME_WORKSHEET );
        }
        $xml .= self::END_CONTENT_TYPES;

        return $xml;
    }

    private function makePackageLevelRelationship()
    {
        $xml = self::START_XML;
        $xml.= self::RELS_START;
        $xml.= str_replace('{workbookPath}', self::WORKBOOK_XML_PATH.'workbook.xml', self::REL_TO_WB);
        $xml.= self::RELS_END;

        return $xml;
    }

    private function makeWorkbook()
    {
        $xml = self::START_XML;
        $xml.= self::WB_START;
        $xml.= self::WB_SHEETS_START;

        $sheetId = 1;
        foreach ($this->worksheets as &$ws)
        {
            $name = is_null($ws->getName()) ? $sheetId : $ws->getName();

            $wsXml = str_replace('{sheetName}', $name, self::WB_SHEET);
            $wsXml = str_replace('{sheetId}', $sheetId, $wsXml);
            $wsXml = str_replace('{rId}', "rId$sheetId", $wsXml);

            $xml.=$wsXml;

            $sheetId++;
        }

        $xml.= self::WB_SHEETS_END;
        $xml.= self::WB_END;
        return $xml;
    }

    private function makeWorkbookRelationship()
    {
        $xml = self::START_XML;
        $xml.= self::RELS_START;

        $rId = 1;
        // Sheets
        for ($sheetId=1; $sheetId<=count($this->worksheets); $sheetId++)
        {
            $sheetRelXml = str_replace('{rId}', "rId$rId", self::WB_RELS_TO_WS);
            $sheetRelXml = str_replace('{sheetPath}', self::WORKSHEET_XML_PATH."sheet$sheetId.xml", $sheetRelXml);
            $xml .= $sheetRelXml;
            $rId++;
        }
        // Styles
        $styleXml = str_replace('{rId}', "rId$rId", self::WB_RELS_TO_STYLES);
        $xml .= str_replace('{stylePath}', self::STYLES_XML_PATH.'styles.xml', $styleXml);


        $xml.= self::RELS_END;

        return $xml;
    }

    private function makeStyles()
    {
        $xml = self::START_XML;
        $xml.= self::STYLES_START;

        $stylesCount = count($this->styles);

        $fonts = str_replace('{count}', $stylesCount, self::STYLES_FONTS_START);
        $colors = self::STYLES_COLORS_START;
        $colors.=self::STYLES_COLOR_INDEXED_START;
        $borders = str_replace('{count}', $stylesCount, self::STYLES_BORDERS_START);

        $fillsOffset=2;
        $fills = str_replace('{count}', $stylesCount+$fillsOffset, self::STYLES_FILLS_START);
        $fills.= self::STYLES_FILLS_EXCEL_DEFAULT;

        $xfs = str_replace('{count}', $stylesCount, self::STYLES_XFS_CELLS_START);
        foreach ($this->styles as &$style)
        {
            $styleId = $style->getId();

            // Fonts
            $fontStyles = $style->getFontStyle();

            $fonts .= self::STYLES_FONT_START;

            if ($fontStyles & SimpleXLSXStyle::FONT_STYLE_BOLD)
            {
                $fonts .= self::STYLES_FONT_STYLE_BOLD;
            }
            if ($fontStyles & SimpleXLSXStyle::FONT_STYLE_ITALIC)
            {
                $fonts .= self::STYLES_FONT_STYLE_ITALIC;
            }
            $fonts .= str_replace('{size}',$style->getFontSize(),self::STYLES_FONT_SIZE);
            $fonts .= str_replace('{colorIndex}',$styleId,self::STYLES_FONT_COLOR);
            $fonts .= str_replace('{name}',$style->getFontName(),self::STYLES_FONT_NAME);
            $fonts .= str_replace('{family}',$style->getFontFamily(),self::STYLES_FONT_FAMILY);

            $fonts .= self::STYLES_FONT_END;

            // Colors
            $colors .= str_replace('{rgbColor}',
                sprintf("%08X",$style->getColor()),
                self::STYLES_COLOR_RGB);

            // Borders
            $borders.=self::STYLES_BORDER_START;
            $brds = $style->getBorders();
            if ( $brds & SimpleXLSXStyle::BORDER_LEFT_THIN )
            {
                $borders .= str_replace('{style}','thin',self::STYLES_BORDER_LEFT);
            } elseif ($brds & SimpleXLSXStyle::BORDER_LEFT_DOUBLE)
            {
                $borders .= str_replace('{style}','double',self::STYLES_BORDER_LEFT);
            }

            if ( $brds & SimpleXLSXStyle::BORDER_RIGHT_THIN )
            {
                $borders .= str_replace('{style}','thin',self::STYLES_BORDER_RIGHT);
            } elseif ($brds & SimpleXLSXStyle::BORDER_RIGHT_DOUBLE)
            {
                $borders .= str_replace('{style}','double',self::STYLES_BORDER_RIGHT);
            }

            if ( $brds & SimpleXLSXStyle::BORDER_TOP_THIN )
            {
                $borders .= str_replace('{style}','thin',self::STYLES_BORDER_TOP);
            } elseif ($brds & SimpleXLSXStyle::BORDER_TOP_DOUBLE)
            {
                $borders .= str_replace('{style}','double',self::STYLES_BORDER_TOP);
            }

            if ( $brds & SimpleXLSXStyle::BORDER_BOTTOM_THIN )
            {
                $borders .= str_replace('{style}','thin',self::STYLES_BORDER_BOTTOM);
            } elseif ($brds & SimpleXLSXStyle::BORDER_BOTTOM_DOUBLE)
            {
                $borders .= str_replace('{style}','double',self::STYLES_BORDER_BOTTOM);
            }
            $borders.=self::STYLES_BORDER_END;

            // Fills
            $fills.= self::STYLES_FILL_START;
            $fills.= self::STYLES_FILL_PATTERN_START;
            $fills.= str_replace(
                '{rgb}',
                sprintf("%'F8X",$style->getBackgroundColor()),
                self::STYLES_FILL_FGCOLOR
            );
            $fills.= self::STYLES_FILL_PATTERN_END;
            $fills.= self::STYLES_FILL_END;

            // xfs
            $xf = str_replace('{fontId}', $styleId, self::STYLES_XFS_START);
            $xf = str_replace('{borderId}', $styleId, $xf);
            $xf = str_replace('{fillId}', $styleId+$fillsOffset, $xf);

            $ha = $this->alignmentToString($style->getHorizontalAlignment());
            $va = $this->alignmentToString($style->getVerticalAlignment());
            $xfa = str_replace('{horizontal}',$ha,self::STYLES_XFS_ALIGNMENT);
            $xfa = str_replace('{vertical}',$va,$xfa);
            $xfa = str_replace('{isWrapText}',
                ($style->getTextWrap() == true) ? '1':'0',
                $xfa
            );
            $xf.= $xfa.self::STYLES_XFS_END;

            $xfs.=$xf;
        }
        $fonts.=self::STYLES_FONTS_END;
        $colors.=self::STYLES_COLOR_INDEXED_END;
        $colors.=self::STYLES_COLORS_END;
        $borders.=self::STYLES_BORDERS_END;
        $xfs.=self::STYLES_XFS_CELLS_END;
        $fills.=self::STYLES_FILLS_END;

        $xml.=$fonts.$fills.$borders.$xfs.$colors;
        $xml.= self::STYLES_END;

        return $xml;
    }

    private function makeWorksheet($worksheet)
    {
        $xml = self::START_XML;

        $xml.= self::WS_START;

        // Cols
        $cols = $worksheet->getColsWidth();
        if ( count( $cols ) > 0 )
        {
            $xml.= self::WS_COLS_START;
            foreach ($cols as $colNum => &$col)
            {
                $colXml = str_replace('{colNum}',$colNum,self::WS_COL);
                $colXml = str_replace('{width}',$col,$colXml);
                $xml.=$colXml;
            }
            $xml.= self::WS_COLS_END;
        }

        // Cells
        $xml.=self::WS_DATA_START;

        $cells = $worksheet->getCells();
        foreach ($cells as $rowNum => &$row)
        {
            $xml.=str_replace('{rowNum}', $rowNum, self::WS_ROW_START);

            foreach ($row as $colNum => &$cell)
            {
                $cellName = $this->cellName($rowNum, $colNum);
                $styleNum = $this->styles[ $cell['style'] ]->getId();

                $xml .= str_replace('{styleNum}',$styleNum,
                    str_replace('{cellName}', $cellName,
                        self::WS_CELL_START)
                );

                $xml .= str_replace('{cellVal}', htmlspecialchars($cell['value']), self::WS_CELL_VALUE);

                $xml.=self::WS_CELL_END;
            }

            $xml.=self::WS_ROW_END;
        }
        $xml.=self::WS_DATA_END;

        $merges = $worksheet->getMerges();
        if ( count($merges)>0 )
        {
            $xml.=self::WS_MERGES_START;

            foreach ($merges as $merge)
            {
                $startCell = $this->cellName($merge['startRow'], $merge['startCol']);
                $endCell = $this->cellName($merge['endRow'], $merge['endCol']);

                $mergeCell = str_replace('{startCell}', $startCell, self::WS_MERGE);
                $mergeCell = str_replace('{endCell}', $endCell, $mergeCell);

                $xml .= $mergeCell;
            }

            $xml.=self::WS_MERGES_END;
        }

        $xml.=self::WS_END;

        return $xml;
    }

    private function alignmentToString($alignment)
    {
        switch ($alignment)
        {
            case SimpleXLSXStyle::ALIGNMENT_GENERAL:
                return 'center';
            case SimpleXLSXStyle::ALIGNMENT_LEFT:
                return 'left';
            case SimpleXLSXStyle::ALIGNMENT_RIGHT:
                return 'right';
            case SimpleXLSXStyle::ALIGNMENT_CENTER:
                return 'center';
            case SimpleXLSXStyle::ALIGNMENT_TOP:
                return 'top';
            case SimpleXLSXStyle::ALIGNMENT_BOTTOM:
                return 'bottom';
        }
        return 'center';
    }

    private function colNumberToColName($colNumber)
    {
        $c = $colNumber;
        $colName='';
        do {
            $pc = intval(($c-1)%26);
            $colName = Chr(Ord('A')+$pc).$colName;
            $c=floor(($c-$pc)/26);
        } while ($c>0);
        return $colName;
    }

    private function cellName($rowNumber, $colNumber)
    {
        $colName = $this->colNumberToColName($colNumber);
        $cellName = $colName.$rowNumber;
        return $cellName;
    }

}