import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;


public class EstilosExcel {

    public static final SimpleDateFormat TIME_FORMAT = new SimpleDateFormat("HH:mm:ss");

    public static CellStyle getConditionalCellStyle(XSSFWorkbook workbook, String menorHora) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();

        try {
            // Converte string "HH:mm:ss" para um objeto Date
            Date timeValue = TIME_FORMAT.parse(menorHora);
            Date limiteVerde = TIME_FORMAT.parse("09:05:59");
            Date limiteAmarelo = TIME_FORMAT.parse("09:10:00");

            if (timeValue.compareTo(limiteVerde) >= 0 && timeValue.compareTo(limiteAmarelo) < 0) {
                // Fundo amarelo claro, fonte amarela escura
                style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
                font.setColor(IndexedColors.DARK_YELLOW.getIndex());
                
            } else if (timeValue.compareTo(limiteAmarelo) >= 0) {
                // Fundo vermelho, fonte branca
                style.setFillForegroundColor(IndexedColors.RED.getIndex());
                font.setColor(IndexedColors.WHITE.getIndex());
            } else {
                // Fundo verde claro, fonte verde escura
                style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                font.setColor(IndexedColors.DARK_GREEN.getIndex());
            }

            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setFont(font);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        return style;
    }
    
    public static CellStyle getCellStyleWithBorders(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        aplicarBordas(style); // aplica a borda

        return style;
    }    

    public static CellStyle getConditionalColumnStyle(XSSFWorkbook workbook, String valorCelula, String media, boolean isTime) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();  

        try {
            boolean isMenor = false;
    
            if (isTime && valorCelula.matches("\\d{2}:\\d{2}:\\d{2}") && media.matches("\\d{2}:\\d{2}:\\d{2}")) {
                // Comparação para tempo no formato HH:mm:ss
                Date total = TIME_FORMAT.parse(valorCelula);
                Date mediaValor = TIME_FORMAT.parse(media);
                //System.out.println(" IF - Comparando: total = " + total + " media = " + mediaValor);                
                isMenor = total.compareTo(mediaValor) < 0;
            } else {
                // Comparação para quantidade numérica
                double total = Double.parseDouble(valorCelula.trim());
                double mediaValor = Double.parseDouble(media.trim());
                //System.out.println(" ELSE - Comparando: total = " + total + " media = " + mediaValor);
                isMenor = total <= mediaValor;
            }
    
            if (isMenor) {
                if (isTime) {
                    
                    // Tempo menor que a média → Amarelo
                    style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
                } else {
                    // Quantidade menor que a média → Laranja
                    style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
                }
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
    
            style.setFont(font);
            font.setBold(true);
            font.setColor(IndexedColors.BLACK.getIndex());
            style.setAlignment(HorizontalAlignment.CENTER);

           aplicarBordas(style); // muda borda quantidade
            
        } catch (ParseException | NumberFormatException e) {
            e.printStackTrace();
        }
    
        return style;
    }
    
    // Estilo do cabeçalho
    public static CellStyle getHeaderCellStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //style.setBorderRight(BorderStyle.DASHED);
        //style.setBorderLeft(BorderStyle.DASHED);
        //style.setBorderTop(BorderStyle.DASHED);
        //style.setBorderBottom(BorderStyle.DASHED);
        aplicarBordas(style);       
        return style;
    }

    // Estilo para a linha de totais
    public static CellStyle getTotalCellStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //style = EstilosExcel.getCellStyleWithBorders(workbook);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.DASHED);
        style.setBorderLeft(BorderStyle.DASHED);
        style.setBorderTop(BorderStyle.DASHED);
        style.setBorderBottom(BorderStyle.DASHED);
        
        return style;
    }

    public static CellStyle aplicarBordas(CellStyle cellStyle){
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        //comentar a baixo depois

        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

        return cellStyle;
    }

     /**
     * Create a library of cell styles
     */
    public static Map<String, CellStyle> criarEstilos(Workbook wb){
        Map<String, CellStyle> estilos = new HashMap<>();
        CellStyle estilo;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)18);
        titleFont.setBold(true);
        estilo = wb.createCellStyle();
        estilo.setAlignment(HorizontalAlignment.CENTER);
        estilo.setVerticalAlignment(VerticalAlignment.CENTER);
        estilo.setFont(titleFont);              

        estilo = wb.createCellStyle();
        estilo.setAlignment(HorizontalAlignment.CENTER);
        estilo.setWrapText(true);
        estilo.setBorderRight(BorderStyle.THIN);
        estilo.setRightBorderColor(IndexedColors.BLACK.getIndex());
        estilo.setBorderLeft(BorderStyle.THIN);
        estilo.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        estilo.setBorderTop(BorderStyle.THIN);
        estilo.setTopBorderColor(IndexedColors.BLACK.getIndex());
        estilo.setBorderBottom(BorderStyle.THIN);
        estilo.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        estilos.put("cell", estilo);

        estilo = wb.createCellStyle();
        estilo.setAlignment(HorizontalAlignment.CENTER);
        estilo.setVerticalAlignment(VerticalAlignment.CENTER);
        estilo.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estilo.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        estilos.put("Media", estilo);

        estilo = wb.createCellStyle();
        estilo.setAlignment(HorizontalAlignment.CENTER);
        estilo.setVerticalAlignment(VerticalAlignment.CENTER);
        estilo.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        estilo.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        estilos.put("Total", estilo);

        return estilos;
    }

    
}
