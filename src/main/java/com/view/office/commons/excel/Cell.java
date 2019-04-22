package com.view.office.commons.excel;

public class Cell {

    private Integer x;

    private Integer y;

    private String text;

    private short borderLeft;

    private short borderLeftColor;

    private short borderRight;

    private short borderRightColor;

    private short borderTop;

    private short borderTopColor;

    private short borderBottom;

    private short borderBottomColor;

    private short align;

    private short bold;

    private short color;

    private short fontHeightInPoints;

    private String background;

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public short getBold() {
        return bold;
    }

    public void setBold(short bold) {
        this.bold = bold;
    }

    public short getColor() {
        return color;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public short getFontHeightInPoints() {
        return fontHeightInPoints;
    }

    public void setFontHeightInPoints(short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }

    public Integer getX() {
        return x;
    }

    public void setX(Integer x) {
        this.x = x;
    }

    public Integer getY() {
        return y;
    }

    public void setY(Integer y) {
        this.y = y;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public short getBorderLeft() {
        return borderLeft;
    }

    public void setBorderLeft(short borderLeft) {
        this.borderLeft = borderLeft;
    }

    public short getBorderRight() {
        return borderRight;
    }

    public void setBorderRight(short borderRight) {
        this.borderRight = borderRight;
    }

    public short getBorderTop() {
        return borderTop;
    }

    public void setBorderTop(short borderTop) {
        this.borderTop = borderTop;
    }

    public short getBorderBottom() {
        return borderBottom;
    }

    public void setBorderBottom(short borderBottom) {
        this.borderBottom = borderBottom;
    }

    public short getBorderLeftColor() {
        return borderLeftColor;
    }

    public void setBorderLeftColor(short borderLeftColor) {
        this.borderLeftColor = borderLeftColor;
    }

    public short getBorderRightColor() {
        return borderRightColor;
    }

    public void setBorderRightColor(short borderRightColor) {
        this.borderRightColor = borderRightColor;
    }

    public short getBorderTopColor() {
        return borderTopColor;
    }

    public void setBorderTopColor(short borderTopColor) {
        this.borderTopColor = borderTopColor;
    }

    public short getBorderBottomColor() {
        return borderBottomColor;
    }

    public void setBorderBottomColor(short borderBottomColor) {
        this.borderBottomColor = borderBottomColor;
    }

    public short getAlign() {
        return align;
    }

    public void setAlign(short align) {
        this.align = align;
    }
}
