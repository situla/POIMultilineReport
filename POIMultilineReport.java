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


import java.util.*;
import java.text.*;
import java.io.*;
import java.nio.charset.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;

public class POIMultilineReport {

    String property_file = "";
    String charset_name = "UTF-8";
    Map<Integer, CellConfig> column_cell_config = new LinkedHashMap<>();
    Integer column_count;

    private int n;
    private StringTokenizer st;

    private String line = null;
    private short rownum;
    private short cellnum;

   /**
    * Method returns the value of a property by some key
    *
    * @param keyname key that corresponds to some value
    * @param fname name of file with properties
    *
    * @return some value or "err" string
    */
    String getValueFromPropertyFile2(String keyname, String fname) {

        Properties prop = new Properties();
        InputStream input = null;
        String ret_val = "err";

        try {
            input = getClass().getClassLoader().getResourceAsStream(fname);
            if (input == null) {
                System.out.println("Sorry, unable to find " + fname);
                return "err";
            }
            prop.load(input);
            Enumeration<?> e = prop.propertyNames();
            while (e.hasMoreElements()) {
                String key = (String) e.nextElement();
                String value = prop.getProperty(key);
                //System.out.println("Key : " + key + ", Value : " + value);
                if (key.equals(keyname)) ret_val = value;
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return ret_val;
    }


    String getValueFromPropertyFile(String keyname, String fname) {
        String value = "err";
        String line = null;

        try {

            BufferedReader in = new BufferedReader(new InputStreamReader(
                                new FileInputStream(fname)/*, Charset.forName("CP1251")*/));

                // search in fname file 
                while ((line = in.readLine()) != null) {
                    if (line.trim().length() == 0) continue;
                    if (line.substring(0, 1).equals("#")) continue;

                    // delimeter is "="
                    StringTokenizer st = new StringTokenizer(line, "=");
                    if ( keyname.equals( st.nextToken() ) )
                        value =  st.hasMoreTokens() ? st.nextToken() : "err";

                }

                in.close();


        } catch (FileNotFoundException fnfe) {
                        System.out.println("WARNING!!! Property file " + fname + " not found!");
        }

        catch (IOException ioe) {
            System.out.println("Error reading file!");
        }


        return value;    
    }
    
    String getFilenamePart(String input) {
        if (input.contains(".")) {
           String[] parts = input.split("\\.");
           return parts[0];
        }  else {
            throw new IllegalArgumentException("String " + input + " does not contain .");
            //return input;
        }
        
    }

    // POIMultilineReport constructor 
    POIMultilineReport (String fileName, String prop_file) throws Exception {

        // it's the property file name field
        this.property_file = prop_file;

        // get CharsetName property
        if (!getValueFromPropertyFile("CharsetName", property_file).equals("err"))
            charset_name = getValueFromPropertyFile("CharsetName", property_file);

        // it's cell configurator
        CellConfig cell_config = null;
        
        // CELLS CONFIGURATOR

        // get column count
        if (!getValueFromPropertyFile("ColumnCount", property_file).equals("err"))
            column_count = Integer.valueOf(getValueFromPropertyFile("ColumnCount", property_file));
        else column_count = 0;

        // fill cell configuration map
        for (Integer i = -2; i < column_count; i++) {
        //for (Integer i = -1; i < column_count; i++) {

            cell_config = new CellConfig();

            if (!getValueFromPropertyFile("FontName"+i.toString(), property_file).equals("err"))
                cell_config.setFontName(getValueFromPropertyFile("FontName"+i.toString(), property_file));
            
            if (!getValueFromPropertyFile("FontSize"+i.toString(), property_file).equals("err"))
                cell_config.setFontSize(Short.valueOf(getValueFromPropertyFile("FontSize"+i.toString(), property_file)));

            if (!getValueFromPropertyFile("FontBold"+i.toString(), property_file).equals("err"))
                cell_config.setFontBold(Boolean.valueOf(getValueFromPropertyFile("FontBold"+i.toString(), property_file)));

            if (!getValueFromPropertyFile("HorizontalAlignment"+i.toString(), property_file).equals("err"))
                cell_config.setHorizontalAlignment(getValueFromPropertyFile("HorizontalAlignment"+i.toString(), property_file));

            if (!getValueFromPropertyFile("ColumnWidth"+i.toString(), property_file).equals("err"))
                cell_config.setColumnWidth(Short.valueOf(getValueFromPropertyFile("ColumnWidth"+i.toString(), property_file)));

            if (!getValueFromPropertyFile("BorderTop"+i.toString(), property_file).equals("err"))
                cell_config.setBorderTop(getValueFromPropertyFile("BorderTop"+i.toString(), property_file));

            if (!getValueFromPropertyFile("BorderLeft"+i.toString(), property_file).equals("err"))
                cell_config.setBorderLeft(getValueFromPropertyFile("BorderLeft"+i.toString(), property_file));

            if (!getValueFromPropertyFile("BorderRight"+i.toString(), property_file).equals("err"))
                cell_config.setBorderRight(getValueFromPropertyFile("BorderRight"+i.toString(), property_file));

            if (!getValueFromPropertyFile("BorderBottom"+i.toString(), property_file).equals("err"))
                cell_config.setBorderBottom(getValueFromPropertyFile("BorderBottom"+i.toString(), property_file));

            //System.out.println(getValueFromPropertyFile("ColumnWidth"+i.toString(), property_file));
            column_cell_config.put(i, cell_config);

        }
    
    FileOutputStream out = new FileOutputStream(getFilenamePart(fileName) + ".xls");
    //FileOutputStream out = new FileOutputStream(fileName + ".xls");
    HSSFWorkbook wb = new HSSFWorkbook();
    HSSFSheet s = wb.createSheet();
        // portrait or landscape orientation
        if (!getValueFromPropertyFile("Landscape", property_file).equals("err"))
            s.getPrintSetup().setLandscape(Boolean.valueOf(getValueFromPropertyFile("Landscape", property_file)));
        else    s.getPrintSetup().setLandscape(false);
    HSSFRow r = null;
    HSSFCell c = null;
    //HSSFCellStyle cs = wb.createCellStyle();
    HSSFCellStyle title_cell_style = wb.createCellStyle();
    HSSFCellStyle header_cell_style = wb.createCellStyle();
    HSSFCellStyle cs_column = null;// = wb.createCellStyle();
    HSSFFont font_column = null;//wb.createFont();
        
        // Create style objects for columns 
        List<HSSFCellStyle> column_cell_style = new ArrayList<HSSFCellStyle>();

        for (Integer i = 0; i < column_count; i++) {
            
            cell_config = new CellConfig();

            if (column_cell_config.containsKey(i)) {
                cell_config.setFontName(column_cell_config.get(i).getFontName());
                cell_config.setFontSize(column_cell_config.get(i).getFontSize());
                cell_config.setFontBold(column_cell_config.get(i).getFontBold());
                cell_config.setHorizontalAlignment(column_cell_config.get(i).getHorizontalAlignment());
                // it's not necessary to put here Column width (only to print)
                cell_config.setColumnWidth(column_cell_config.get(i).getColumnWidth());
                cell_config.setBorderTop(column_cell_config.get(i).getBorderTop());
                cell_config.setBorderLeft(column_cell_config.get(i).getBorderLeft());
                cell_config.setBorderRight(column_cell_config.get(i).getBorderRight());
                cell_config.setBorderBottom(column_cell_config.get(i).getBorderBottom());
            }

            System.out.println("Column " + i.toString() + "\n--------------\n" + cell_config.toString()  + "\n--------------\n");

            font_column = wb.createFont();

            // font size
            font_column.setFontHeightInPoints((short)cell_config.getFontSize());

            // font name
            font_column.setFontName(cell_config.getFontName());

            // set font bold or not
            font_column.setBold(cell_config.getFontBold());

            // create style object
            cs_column = wb.createCellStyle();

            cs_column.setAlignment(HorizontalAlignment.valueOf(cell_config.getHorizontalAlignment()));
            cs_column.setVerticalAlignment(VerticalAlignment.valueOf("TOP"));

            cs_column.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));

            cs_column.setBorderTop(BorderStyle.valueOf(cell_config.getBorderTop()));
            cs_column.setBorderLeft(BorderStyle.valueOf(cell_config.getBorderLeft()));
            cs_column.setBorderRight(BorderStyle.valueOf(cell_config.getBorderRight()));
            cs_column.setBorderBottom(BorderStyle.valueOf(cell_config.getBorderBottom()));
            
            cs_column.setFont(font_column);

            // https://poi.apache.org/apidocs/org/apache/poi/hssf/usermodel/HSSFCellStyle.html#setWrapText-boolean-            
            cs_column.setWrapText(true);

            column_cell_style.add(cs_column);

        }

        // Create style objects for even columns 
        List<HSSFCellStyle> column_cell_style_even = new ArrayList<HSSFCellStyle>();

        for (Integer i = 0; i < column_count; i++) {
            
            cell_config = new CellConfig();

            if (column_cell_config.containsKey(i)) {
                cell_config.setFontName(column_cell_config.get(i).getFontName());
                cell_config.setFontSize(column_cell_config.get(i).getFontSize());
                cell_config.setFontBold(column_cell_config.get(i).getFontBold());
                cell_config.setHorizontalAlignment(column_cell_config.get(i).getHorizontalAlignment());
                // it's not necessary to put here Column width (only to print)
                cell_config.setColumnWidth(column_cell_config.get(i).getColumnWidth());
                cell_config.setBorderTop(column_cell_config.get(i).getBorderTop());
                cell_config.setBorderLeft(column_cell_config.get(i).getBorderLeft());
                cell_config.setBorderRight(column_cell_config.get(i).getBorderRight());
                cell_config.setBorderBottom(column_cell_config.get(i).getBorderBottom());
            }

            System.out.println("Column " + i.toString() + "\n--------------\n" + cell_config.toString()  + "\n--------------\n");

            font_column = wb.createFont();

            // font size
            font_column.setFontHeightInPoints((short)cell_config.getFontSize());

            // font name
            font_column.setFontName(cell_config.getFontName());

            // set font bold or not
            font_column.setBold(cell_config.getFontBold());

            // create style object
            cs_column = wb.createCellStyle();

            cs_column.setAlignment(HorizontalAlignment.valueOf(cell_config.getHorizontalAlignment()));
            cs_column.setVerticalAlignment(VerticalAlignment.valueOf("TOP"));

            cs_column.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));

            cs_column.setBorderTop(BorderStyle.valueOf(cell_config.getBorderTop()));
            cs_column.setBorderLeft(BorderStyle.valueOf(cell_config.getBorderLeft()));
            cs_column.setBorderRight(BorderStyle.valueOf(cell_config.getBorderRight()));
            cs_column.setBorderBottom(BorderStyle.valueOf(cell_config.getBorderBottom()));
            
            cs_column.setFont(font_column);

            // https://poi.apache.org/apidocs/org/apache/poi/hssf/usermodel/HSSFCellStyle.html#setWrapText-boolean-            
            cs_column.setWrapText(true);

            cs_column.setFillPattern(FillPatternType.valueOf("SOLID_FOREGROUND"));
            cs_column.setFillForegroundColor(HSSFColor.HSSFColorPredefined.valueOf("PALE_BLUE").getIndex());

            column_cell_style_even.add(cs_column);

        }

        // this font for Report Title
        HSSFFont title_font = wb.createFont();
        
        // get style for report title
        cell_config = new CellConfig();
        
        if (column_cell_config.containsKey(-1)) {
            cell_config.setFontName(column_cell_config.get(-1).getFontName());
            cell_config.setFontSize(column_cell_config.get(-1).getFontSize());
            cell_config.setFontBold(column_cell_config.get(-1).getFontBold());
        }

        title_font.setFontHeightInPoints((short)cell_config.getFontSize());

        title_font.setFontName(cell_config.getFontName());

        title_font.setBold(cell_config.getFontBold());

    
        // title cell style
        title_cell_style.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));
        // set font
        title_cell_style.setFont(title_font);

        // sheet name
        wb.setSheetName(0, "Sheet Name" );

        BufferedReader in = new BufferedReader(new InputStreamReader(
            new FileInputStream(fileName)/*, Charset.forName("CP1251")*/));

        rownum = (short) 0;
    
        // cell for title
        r = s.createRow(rownum);
        cellnum = (short) 0;
        c = r.createCell(cellnum);
        // set height of cell
        r.setHeight((short) 450);
        // set cell stile
        c.setCellStyle(title_cell_style);
        // set Document Title
        if (getValueFromPropertyFile("DocumentTitle", property_file).equals("err"))
            c.setCellValue("Set the DocumentTitle property in " + property_file + " file!");
        else {
            String val;
            val = getValueFromPropertyFile("DocumentTitle", property_file);
            
            if (getValueFromPropertyFile("AddSourceFileName", property_file).equals("true")) {
            
                val += " " + getFilenamePart(fileName);
            }
                
            if (getValueFromPropertyFile("AddCurrentDate", property_file).equals("true")) {
                
                Date dateNow = new Date();
                SimpleDateFormat formatForDateNow = new SimpleDateFormat("dd.MM.yyyy");
                           
                val += " (" + formatForDateNow.format(dateNow) + ")";
            }    
            
            c.setCellValue(val);
        }
    
        rownum++;
        
        // this font for table header
        HSSFFont header_font = wb.createFont();

        // get style for table header
        cell_config = new CellConfig();

        if (column_cell_config.containsKey(-2)) {
            cell_config.setFontName(column_cell_config.get(-2).getFontName());
            cell_config.setFontSize(column_cell_config.get(-2).getFontSize());
            cell_config.setFontBold(column_cell_config.get(-2).getFontBold());
            cell_config.setHorizontalAlignment(column_cell_config.get(-2).getHorizontalAlignment());
            cell_config.setBorderTop(column_cell_config.get(-2).getBorderTop());
            cell_config.setBorderLeft(column_cell_config.get(-2).getBorderLeft());
            cell_config.setBorderRight(column_cell_config.get(-2).getBorderRight());
            cell_config.setBorderBottom(column_cell_config.get(-2).getBorderBottom());

        }

        header_font.setFontHeightInPoints((short)cell_config.getFontSize());

        header_font.setFontName(cell_config.getFontName());

        header_font.setBold(cell_config.getFontBold());

        // title cell style
        header_cell_style.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));
        
        header_cell_style.setBorderTop(BorderStyle.valueOf(cell_config.getBorderTop()));
        header_cell_style.setBorderLeft(BorderStyle.valueOf(cell_config.getBorderLeft()));
        header_cell_style.setBorderRight(BorderStyle.valueOf(cell_config.getBorderRight()));
        header_cell_style.setBorderBottom(BorderStyle.valueOf(cell_config.getBorderBottom()));

        // set font
        header_cell_style.setFont(header_font);

        header_cell_style.setAlignment(HorizontalAlignment.valueOf(cell_config.getHorizontalAlignment()));

        r = s.createRow(rownum);

        // table header
        String header_column_value = "";
        if (!getValueFromPropertyFile("Values-2", property_file).equals("err"))
            header_column_value = getValueFromPropertyFile("Values-2", property_file);
        st = new StringTokenizer(header_column_value, "$");

        for (int i = 0; i < column_count; i++) {
            // set height of cell
            r.setHeight((short) 350);
            cellnum = (short) i;
            c = r.createCell(cellnum);
            c.setCellStyle(header_cell_style);
            if (st.hasMoreTokens()) c.setCellValue(st.nextToken());
        }
    
        rownum++;
    
        int print_row_num = 1;

    while( (line = in.readLine()) != null) {
      if(line.trim().length()==0) break;
      
      r = s.createRow(rownum);
      //r.setHeight((short) 400);
      // this is 'auto' height see setWrapText(true);
      r.setHeight((short)-1);
      
      st = new StringTokenizer(line, "#");
      
      n = st.countTokens();
      String[] a = new String[n];
      
      for (int j = 0; j < n; j++) {
        
        a[j] = st.nextToken();
        cellnum = (short) j;
            // create new cell
            c = r.createCell(cellnum);
            
            if (print_row_num%2 == 0) {
                c.setCellStyle(column_cell_style.get(j));    
            } else c.setCellStyle(column_cell_style_even.get(j));


            //s.setColumnWidth((short) cellnum, (short)5000);
            s.setColumnWidth((short) cellnum, column_cell_config.get(j).getColumnWidth());

            // TODO formula or text ...
        
        c.setCellValue(a[j]);
        
      }
      rownum++;
      print_row_num++;
     
    }
    
    in.close();
    
    wb.write(out);
    out.close();
        
    return;
  }

    public static void main (String args[]) throws Exception {
        String csv_file = "", property_file = "";
        int good_args_count = 0;

        for (int n = 0; n < args.length; n++) {
            if (args[n].equals("-f")) {
                csv_file = args[++n];
                good_args_count++;
            }
        }

        for (int n = 0; n < args.length; n++) {
            if (args[n].equals("-p")) {
                property_file = args[++n];
                good_args_count++;
            }
        }
        if (good_args_count != 2) throw new IllegalArgumentException("Wrong arguments!\nUsage must be:\n'java ru.learn2prog.poi.POIMultilineReport -f filename.csv -p property.file'\nwhere property.file is a file with main autoreplace rules");
                
        new POIMultilineReport (csv_file, property_file);
        
    }

}
