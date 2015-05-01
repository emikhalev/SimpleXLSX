<?php
/**
 * Created by IntelliJ IDEA.
 * User: eugeny.mikhalev
 * Date: 07.07.14
 * Time: 15:03
 * To change this template use File | Settings | File Templates.
 */
namespace emikhalev\SimpleXLSX;

class SimpleXLSXStyle
{
    const FONT_STYLE_NORMAL = 0;
    const FONT_STYLE_BOLD   = 1;
    const FONT_STYLE_ITALIC = 2;

    const COLOR_AUTO = null;
    const COLOR_TEXT_DEFAULT = 0x0;
    const COLOR_BACKGROUND_DEFAULT = 0x00FFFFFF;

    const BORDER_NONE           = 0;
    const BORDER_LEFT_THIN      = 1;
    const BORDER_RIGHT_THIN     = 2;
    const BORDER_TOP_THIN       = 4;
    const BORDER_BOTTOM_THIN    = 8;
    const BORDER_LEFT_DOUBLE    = 16;
    const BORDER_RIGHT_DOUBLE   = 32;
    const BORDER_TOP_DOUBLE     = 64;
    const BORDER_BOTTOM_DOUBLE  = 128;

    const ALIGNMENT_GENERAL = 0;
    const ALIGNMENT_CENTER  = 1;
    const ALIGNMENT_LEFT    = 2;
    const ALIGNMENT_RIGHT   = 3;
    const ALIGNMENT_TOP     = 4;
    const ALIGNMENT_BOTTOM  = 5;

    private $id = null;
    private $fontStyle = self::FONT_STYLE_NORMAL;
    private $color = self::COLOR_TEXT_DEFAULT;
    private $bgColor = self::COLOR_BACKGROUND_DEFAULT;
    private $borders = self::BORDER_NONE;
    private $fontName = 'Arial';
    private $fontFamily = 2;
    private $fontSize = 12;
    private $horizontalAlignment = self::ALIGNMENT_GENERAL;
    private $verticalAlignment = self::ALIGNMENT_GENERAL;
    private $textWrap = false;

    public function __construct($styleId)
    {
        $this->id = $styleId;
    }

    public function getId()
    {
        return $this->id;
    }

    public function setFontName($fontName)
    {
        $this->fontName = $fontName;
    }

    public function getFontName()
    {
        return $this->fontName;
    }

    public function setFontFamily($fontFamily)
    {
        $this->fontFamily = $fontFamily;
    }

    public function getFontFamily()
    {
        return $this->fontFamily;
    }

    public function setFontStyle($style = self::FONT_STYLE_NORMAL)
    {
        $this->fontStyle = $style;
    }

    public function getFontStyle()
    {
        return $this->fontStyle;
    }

    public function setFontSize($fontSize=12)
    {
        $this->fontSize = $fontSize;
    }

    public function getFontSize()
    {
        return $this->fontSize;
    }

    public function setColor($color = self::COLOR_TEXT_DEFAULT)
    {
        $this->color = $color;
    }

    public function getColor()
    {
        return $this->color;
    }

    public function setBackgroundColor($bgColor = self::COLOR_BACKGROUND_DEFAULT)
    {
        $this->bgColor = $bgColor;
    }

    public function getBackgroundColor()
    {
        return $this->bgColor;
    }

    public function setBorders($borders = self::BORDER_NONE)
    {
        $this->borders = $borders;
    }

    public function getBorders()
    {
        return $this->borders;
    }

    public function setHorizontalAlignment($alignment)
    {
        $this->horizontalAlignment = $alignment;
    }

    public function getHorizontalAlignment()
    {
        return $this->horizontalAlignment;
    }

    public function setVerticalAlignment($alignment)
    {
        $this->verticalAlignment = $alignment;
    }

    public function getVerticalAlignment()
    {
        return $this->verticalAlignment;
    }

    public function getTextWrap()
    {
        return $this->textWrap;
    }

    public function setTextWrap($isWrapText = false)
    {
        $this->textWrap = $isWrapText;
    }
}