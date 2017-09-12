package com.liqiaoqi.service;

import com.liqiaoqi.util.Typhoon;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * Created by LQQ on 2017/9/11 0011.
 */
@Service
public class ReadExelService {

    private List<Typhoon> typhoonsList = new ArrayList<Typhoon>();  //用一个list保存计算好的Typhoon对象

    public void readTyphoonInfo() throws Exception{
        //Excel文件
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(ResourceUtils.getFile("classpath:台风信息.xlsm")));
        //台风信息工作表
        XSSFSheet sheet = wb.getSheetAt(0);

        //循环读出每条记录，并进行计算
        for(int i = 200; i <= sheet.getLastRowNum();i++ ){
            XSSFRow row = sheet.getRow(i);
            XSSFCell typhooncodeCell = row.getCell(0);
            typhooncodeCell.setCellType(XSSFCell.CELL_TYPE_STRING); //设置单元格格式，否则无法用字符串接收一个将数字格式的cell
            String typhooncode = typhooncodeCell.getStringCellValue();
            Date date = row.getCell(2).getDateCellValue();
            Date date1 = row.getCell(3).getDateCellValue();

            int days = calculateDaysBetween(date,date1);//计算台风天数
            double rainFall = calculateRainFall(days,date,date1); //计算面雨量
            Typhoon typhoon = new Typhoon();
            typhoon.setTyphooncode(typhooncode);
            typhoon.setRainFall(rainFall);
            typhoonsList.add(typhoon);
            System.out.println("台风码："+typhooncode+"\t相差天数是："+ days +"\t面雨量是:" + rainFall );
        }
        writeExel();

    }

    /**
     * 计算两个日期之间相差的具体天数
     * @param firstDate
     * @param seconddate
     * @return
     */
    public int calculateDaysBetween(Date firstDate,Date seconddate){
        //初始化两个Calendar实例
        Calendar firstCalendar = Calendar.getInstance();
        Calendar secondCalendar = Calendar.getInstance();
        //将date对象转换为Calendar类型
        firstCalendar.setTime(firstDate);
        secondCalendar.setTime(seconddate);
        //将时、分、秒置零
        firstCalendar.set(Calendar.HOUR_OF_DAY, 0);
        firstCalendar.set(Calendar.MINUTE, 0);
        firstCalendar.set(Calendar.SECOND, 0);

        secondCalendar.set(Calendar.HOUR_OF_DAY, 0);
        secondCalendar.set(Calendar.MINUTE, 0);
        secondCalendar.set(Calendar.SECOND, 0);
        //计算两个日期相差的天数
        return ((int)(secondCalendar.getTime().getTime()/1000)-(int)(firstCalendar.getTime().getTime()/1000))/3600/24 + 1;
        //加1是因为我需要的天数包括开始和结束的那两天
    }


    /**
     * 计算一场台风的总面雨量
     * @param days
     * @param firstDate
     * @param seconddate
     * @return
     * @throws Exception
     */
    public double calculateRainFall (int days,Date firstDate,Date seconddate) throws Exception{

        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(ResourceUtils.getFile("classpath:台风信息.xlsm")));
        //宁波市面雨量工作表
        XSSFSheet sheet = wb.getSheetAt(4);

        Date date = new Date();
        List<Double> rainFallList = new ArrayList<Double>();

        for(int i = 1 ; i <= sheet.getLastRowNum(); i++){
            XSSFRow row = sheet.getRow(i);
            date = row.getCell(1).getDateCellValue();
            //由于表中数据的特殊性，所以要为date加上一天的毫秒数，以防止有的数据筛选不出来
            if(date.getTime()+ 86400000 - 1 >= firstDate.getTime() && date.getTime() <= seconddate.getTime() ){
                Double rainFall = (Double) row.getCell(2).getNumericCellValue();
                rainFallList.add(rainFall);
                System.out.println(rainFall);
            }
        }

        if(days <= 3 ){  //若天数小于3，则面雨量直接累加
            double sum = 0;
            for(int j = 0; j < rainFallList.size() ; j++){
                sum += rainFallList.get(j);
            }
            return sum;
        }else{
            return maxRainFall(rainFallList);
        }
    }

    /**
     * 当台风日期大于三天时，用特定方法计算其面雨量
     * @param rainFallList
     * @return
     */
    public double maxRainFall(List<Double> rainFallList){
        double maxRainFall = 0.0;
        for(int i = 0 ; i < rainFallList.size() - 2 ; i++){
            double rainFall = rainFallList.get(i) + rainFallList.get(i+1) + rainFallList.get(i+2);
            System.out.println("第"+(i+1)+"次相加："+rainFall);
            if(rainFall > maxRainFall){
                maxRainFall = rainFall;
            }
        }
        return maxRainFall;
    }

    /**
     * 将计算的结果写入exel表
     * @throws Exception
     */
    public void writeExel() throws Exception{
        String excelFileName = "classpath:台风面雨量.xlsm";// 文件名

        String sheetName = "台风面雨量";// 工作表名

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        // 循环写入每条记录的台风码和面雨量
        for (int r = 1; r < typhoonsList.size(); r++)
        {
            XSSFRow row = sheet.createRow(r);

            XSSFCell typhooncodeCell = row.createCell(0);

            typhooncodeCell.setCellValue(typhoonsList.get(r - 1).getTyphooncode());

            XSSFCell rainfallCell1 = row.createCell(1);

            rainfallCell1.setCellValue(typhoonsList.get(r - 1).getRainFall());

        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        // 将workbook写到输出流中
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

}
