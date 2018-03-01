import datos.CoursesJDBC;
import app.Course;
import java.util.List;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.Map;
import java.util.ArrayList;
import java.util.HashMap;
import java.io.FileOutputStream;


public class Reporte {
	
	 private static final String[] titles = {"Id","ShortName","FullName"};
	 
	public static void main(String[] args) throws Exception {
		CoursesJDBC coursesJDBC = new CoursesJDBC();
		List<Course> cursos = coursesJDBC.select();
		
		ArrayList<String> name = new ArrayList<String>();
			
		int inc = 0;
	        for (Course curso : cursos) {
	        	name.add(curso.getFullname());
	        	inc++;
	        }
	        
	        String[] array = name.toArray(new String[0]);
	        String[][] data = new String[array.length][array.length];
	        
	        int ele = array.length; 
	        System.out.println(ele);
	        	
	        		for (int b = 0; b < ele; b++) {
	        			int indice = 0;
	        			for (Course curso : cursos) {
	        			

		        				data[indice][0] = String.valueOf(curso.getId());

		        				data[indice][1] = curso.getShortname();
		        				
		        				data[indice][2] = curso.getFullname();

		        			indice++;
	        			
	        			}
	        			
	        		}
	        		


	        Workbook wb;

	        if(args.length > 0 && args[0].equals("-xls")) wb = new HSSFWorkbook();
	        else wb = new XSSFWorkbook();

	        Map<String, CellStyle> styles = createStyles(wb);

	        Sheet sheet = wb.createSheet("Courses Report");
	        
	        //turn off gridlines
	        sheet.setDisplayGridlines(false);
	        sheet.setPrintGridlines(false);
	        sheet.setFitToPage(true);
	        sheet.setHorizontallyCenter(true);
	        PrintSetup printSetup = sheet.getPrintSetup();
	        printSetup.setLandscape(true);

	        //the following three statements are required only for HSSF
	        sheet.setAutobreaks(true);
	        printSetup.setFitHeight((short)1);
	        printSetup.setFitWidth((short)1);
	        
	      //the header row: centered text in 48pt font
	        Row headerRow = sheet.createRow(0);
	        headerRow.setHeightInPoints(12.75f);
	        for (int i = 0; i < titles.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(titles[i]);
	            cell.setCellStyle(styles.get("header"));
	        }

	        //freeze the first row
	        sheet.createFreezePane(0, 1);
	        
	        Row row;
	        Cell cell;
	        int rownum = 1;
	        for (int i = 0; i < data.length; i++, rownum++) {
	            row = sheet.createRow(rownum);
	            if(data[i] == null) continue;

	            for (int j = 0; j < data[i].length; j++) {
	                cell = row.createCell(j);
	                String styleName;
	                boolean isHeader = i == 0 || data[i-1] == null;
	                switch(j){
	                    case 0:
	                        if(isHeader) {
	                            styleName = "cell_b";
	                            cell.setCellValue(Double.parseDouble(data[i][j]));
	                        } else {
	                            styleName = "cell_normal";
	                            cell.setCellValue(data[i][j]);
	                        }
	                        break;
	                    case 1:
	                        if(isHeader) {
	                            styleName = i == 0 ? "cell_h" : "cell_bb";
	                        } else {
	                            styleName = "cell_indented";
	                        }
	                        cell.setCellValue(data[i][j]);
	                        break;
	                    case 2:
	                        styleName = isHeader ? "cell_b" : "cell_normal";
	                        cell.setCellValue(data[i][j]);
	                        break;
	                    default:
	                        styleName = data[i][j] != null ? "cell_blue" : "cell_normal";
	                }

	                cell.setCellStyle(styles.get(styleName));
	            }
	            
	        }
	        
	        //group rows for each phase, row numbers are 0-based
	        sheet.groupRow(4, 6);
	        sheet.groupRow(9, 13);
	        sheet.groupRow(16, 18);

	        //set column widths, the width is measured in units of 1/256th of a character width
	        sheet.setColumnWidth(0, 256*6);
	        sheet.setColumnWidth(1, 256*33);
	        sheet.setColumnWidth(2, 256*20);
	        sheet.setZoom(75); //75% scale


	        // Write the output to a file
	        String file = "course_report_2.xls";
	        if(wb instanceof XSSFWorkbook) file += "x";
	        FileOutputStream out = new FileOutputStream(file);
	        wb.write(out);
	        out.close();
	        
	        wb.close();
	        System.out.println("Ok. Terminate. Xls Generated...");
	}
	
	/**
	  * create a library of cell styles
	  */
	 private static Map<String, CellStyle> createStyles(Workbook wb){
	     Map<String, CellStyle> styles = new HashMap<>();
	     DataFormat df = wb.createDataFormat();

	     CellStyle style;
	     Font headerFont = wb.createFont();
	     headerFont.setBold(true);
	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.CENTER);
	     style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
	     style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	     style.setFont(headerFont);
	     styles.put("header", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.CENTER);
	     style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
	     style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	     style.setFont(headerFont);
	     style.setDataFormat(df.getFormat("d-mmm"));
	     styles.put("header_date", style);

	     Font font1 = wb.createFont();
	     font1.setBold(true);
	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.LEFT);
	     style.setFont(font1);
	     styles.put("cell_b", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.CENTER);
	     style.setFont(font1);
	     styles.put("cell_b_centered", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.RIGHT);
	     style.setFont(font1);
	     style.setDataFormat(df.getFormat("d-mmm"));
	     styles.put("cell_b_date", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.RIGHT);
	     style.setFont(font1);
	     style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	     style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	     style.setDataFormat(df.getFormat("d-mmm"));
	     styles.put("cell_g", style);

	     Font font2 = wb.createFont();
	     font2.setColor(IndexedColors.BLUE.getIndex());
	     font2.setBold(true);
	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.LEFT);
	     style.setFont(font2);
	     styles.put("cell_bb", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.RIGHT);
	     style.setFont(font1);
	     style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	     style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	     style.setDataFormat(df.getFormat("d-mmm"));
	     styles.put("cell_bg", style);

	     Font font3 = wb.createFont();
	     font3.setFontHeightInPoints((short)14);
	     font3.setColor(IndexedColors.DARK_BLUE.getIndex());
	     font3.setBold(true);
	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.LEFT);
	     style.setFont(font3);
	     style.setWrapText(true);
	     styles.put("cell_h", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.LEFT);
	     style.setWrapText(true);
	     styles.put("cell_normal", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.CENTER);
	     style.setWrapText(true);
	     styles.put("cell_normal_centered", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.RIGHT);
	     style.setWrapText(true);
	     style.setDataFormat(df.getFormat("d-mmm"));
	     styles.put("cell_normal_date", style);

	     style = createBorderedStyle(wb);
	     style.setAlignment(HorizontalAlignment.LEFT);
	     style.setIndention((short)1);
	     style.setWrapText(true);
	     styles.put("cell_indented", style);

	     style = createBorderedStyle(wb);
	     style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
	     style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	     styles.put("cell_blue", style);

	     return styles;
	 }

	 private static CellStyle createBorderedStyle(Workbook wb){
	     BorderStyle thin = BorderStyle.THIN;
	     short black = IndexedColors.BLACK.getIndex();
	     
	     CellStyle style = wb.createCellStyle();
	     style.setBorderRight(thin);
	     style.setRightBorderColor(black);
	     style.setBorderBottom(thin);
	     style.setBottomBorderColor(black);
	     style.setBorderLeft(thin);
	     style.setLeftBorderColor(black);
	     style.setBorderTop(thin);
	     style.setTopBorderColor(black);
	     return style;
	 }
}
