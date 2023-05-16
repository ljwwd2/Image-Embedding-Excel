package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        creatExcelFile();
    }
    public static void creatExcelFile() throws IOException {
        List<String> errList = new ArrayList<>();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("sheet1");
        //1. 设置sheet的默认单元格行高和宽度
        sheet.setDefaultRowHeight((short) (36 * 20));//120
        sheet.setDefaultColumnWidth(20);//70
        //2. 创建首行，写入头部
        final XSSFRow row = sheet.createRow(0);
        for (int i = 0; i < 7; i++) {
            //设置单元格格式
            final XSSFCell cell = row.createCell(i);
            CellStyle style = workbook.createCellStyle();
            //背景颜色
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(CellStyle.BIG_SPOTS);
            style.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
            style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中

            cell.setCellStyle(style);
            row.setHeight((short)(15*20));//设置行高
            cell.setCellValue("第"+(i+1)+"排");//插入值
        }

        //3. 写入内容
        //条码格式为A-BB-CC-DD，由于中间编号CC会变化，所以用数组记录一下
        int numberBB = 7;//编号BB的最大值
        int[] numbersCC = {10,10,12,12,12,12,13};
        int numberDD = 6;//编号DD的最大值

        for (int k = 1; k <= numberBB; k++) {// k 表示文件夹名称 1,2,3,4,5,6,7
            //获取第k个文件夹中CC的值
            int NumberCC = numbersCC[k-1];
            //每个文件夹中的图片从k列的第rowNum行开始写入,每写入一次rowNum++
            int rowNum = 1;
            for (int j = 1; j <= NumberCC; j++) {
                for (int i = 1; i <= numberDD; i++) {
                    // fileName为图片完整路径，例：C:\images\EDG.jpg
                    String picFileName = System.getProperty("user.dir")+ File.separator+"file"+File.separator+k+"\\A-"+getPrefixNumToStringBy2(k)+"-"+getPrefixNumToStringBy2(j)+"-"+getPrefixNumToStringBy2(i)+".png";
                    //生成图片的主要逻辑方法--writePicToExcel()
                    final String errMsg = writePicToExcel(workbook, sheet, k-1, rowNum++, picFileName, Workbook.PICTURE_TYPE_PNG);
                    if(errMsg!=null){//记录错误信息
                        errList.add(errMsg);
                    }
                }
            }
        }
        if (errList.size()!=0){
            System.err.println("下列图片不存在："+errList);
        }
        try (OutputStream fileOut = Files.newOutputStream(Paths.get(System.getProperty("user.dir") + File.separator +System.currentTimeMillis()+ "workbook.xlsx"))) {
            workbook.write(fileOut);
        }
    }

    /**
     * 向excel的单元格中嵌入图片
     * @param workbook org.apache.poi.xssf.usermodel.XSSFWorkbook
     * @param sheet  org.apache.poi.xssf.usermodel.XSSFSheet
     * @param colNum 单元格行号
     * @param rowNum 单元格列号
     * @param picFilePath 图片文件地址  ./A-01-01-01.png
     * @param picType 图片文件类型
     * @return String 错误信息
     */
    public static String writePicToExcel(XSSFWorkbook workbook, XSSFSheet sheet, int colNum, int rowNum, String picFilePath, int picType){
        try (InputStream is = Files.newInputStream(Paths.get(picFilePath))){
            byte[] bytes = IOUtils.toByteArray(is);
            // 这里根据实际需求选择图片类型
            // 参数： 1.图片格式的字节  2.图片的格式。 返回： 此图片的索引（基于0），添加的图片可以从getAllPictures（）获得。
            int pictureIdx = workbook.addPicture(bytes, picType);

            //workbook.getCreationHelper() 返回一个对象，该对象处理XSSF的各种实例的具体类的实例化
            ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();//CreationHelper创建帮助程序
            anchor.setCol1(colNum); // 列号
            anchor.setRow1(rowNum); // 行号

            //设置图片规格，即图片在当前单元格（k-1,rowNum）中的大小
            anchor.setDx1(org.apache.poi.util.Units.EMU_PER_PIXEL);//由官方文档可知，设置第一个单元格中的x坐标注意-XSSF和HSSF的坐标系略有不同，XSSF中的值要大一个系数org.apache.poi.util.Units.EMU_PER_PIXEL
            anchor.setDx2(org.apache.poi.util.Units.EMU_PER_PIXEL);
            anchor.setDy1(org.apache.poi.util.Units.EMU_PER_PIXEL);
            // 插入图片
            Drawing drawing = sheet.createDrawingPatriarch();//创建一个新的电子表格ML图形。如果此sheet已经包含一个图形，请返回该图形。 指定人： 在接口表中创建DrawingPatriarch 返回值： 电子表格ML sheet
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //调整图片占单元格百分比的大小，1.0就是100%
            //pict.resize(1);
            pict.resize(0.995,1);
        }catch (IOException e){
            return e.getMessage();
        }
        return null;
    }

    //获取长度为2的数字字符串
    public static String getPrefixNumToStringBy2(int num) {
        if (num >= 100) {
            return "-1";
        }else if (num>=10){
            return String.valueOf(num);
        }
        return "0"+num;
    }


}