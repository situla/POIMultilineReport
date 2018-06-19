/**
 * Copyright (C) 2018  by situla
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */

package ru.learn2prog.poi;

/**
 * This class describe of Excel cell configuration properties.
 * Contains font properties (name, size, bold/not bold, etc.)
 * You may add another cell configuration properties
 */
 
public class CellConfig {
    
    // Font section
    private String font_name;
    private Short font_size, column_width;
    private Boolean font_bold;

    // border section
    private String border_bottom, border_top, border_left, border_right;

    // Horizontal Alignment section
    private String horizontal_alignment;
    
    // cell_type may be numeric, formula or string.
    // we are use string or formula
    private String cell_type;
    private String formula;

    CellConfig() {

        // default config options
        font_name = "TimesNewRoman";
        font_size = 12;
        font_bold = false;

        border_bottom = "THIN";
        border_top = "THIN";
        border_left = "THIN";
        border_right = "THIN";

        horizontal_alignment = "LEFT";
        column_width = 5000;
        cell_type = "string";
        formula = "";
    }

    public void setFontName(String font_name) {
        this.font_name = font_name;
    }

    public String getFontName() {
        return font_name;
    }

    public void setFontSize(Short font_size) {
        this.font_size = font_size;
    }

    public Short getFontSize() {
        return font_size;
    }

    public void setFontBold(Boolean font_bold) {
        this.font_bold = font_bold;
    }

    public Boolean getFontBold() {
        return font_bold;
    }

    public void setBorderBottom(String border_bottom) {
        this.border_bottom = border_bottom;
    }

    public String getBorderBottom() {
        return border_bottom;
    }

    public void setBorderTop(String border_top) {
        this.border_top = border_top;
    }

    public String getBorderTop() {
        return border_top;
    }

    public void setBorderLeft(String border_left) {
        this.border_left = border_left;
    }

    public String getBorderLeft() {
        return border_left;
    }

    public void setBorderRight(String border_right) {
        this.border_right = border_right;
    }

    public String getBorderRight() {
        return border_right;
    }

    public void setHorizontalAlignment(String horizontal_alignment) {
        this.horizontal_alignment = horizontal_alignment;
    }

    public String getHorizontalAlignment() {
        return horizontal_alignment;
    }
    
    public void setColumnWidth(Short column_width) {
        this.column_width = column_width;
    }
    
    public Short getColumnWidth() {
        return column_width;
    }
    
    public void setCellType(String cell_type) {
        this.cell_type = cell_type;
    }

    public String getCellType() {
        return cell_type;
    }

    public void setFormula (String formula) {
        this.formula = formula;
    }

    public String getFormula() {
        return formula;
    }

    public String toString() {
        String frml = (this.formula.equals("")) ? "\n" : "\n" + this.formula ;
        return "STATE OF CELLS: \n" + "Font: " + font_name + " " + font_size.toString() + " " + font_bold.toString() + "\n" +
                "Borders: " + border_bottom + ", " + border_left + ", " + border_right + ", " + border_top + "\n" +
                "Alignment: " + horizontal_alignment + "\n" +
                "Column width: " + column_width + "\n" +
                "Cell type: " + cell_type +
                frml;

    }

    public static void main(String[] args) {
        CellConfig cc = new CellConfig();
        //System.out.println(cc.border_bottom);
        System.out.println( cc.toString() );
    }


}
