package com.liqiaoqi.service;

import com.liqiaoqi.util.TyphoonInfluence;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by LQQ on 2017/9/12 0012.
 */
@Service
public class TyphoonInfluenceService {
    List<TyphoonInfluence> typhoonInfluenceList = new ArrayList<TyphoonInfluence>(); //保存TyphoonInfluence对象的列表

    /**
     * 主类，读取每场台风对应的台风码和面雨量
     * @throws Exception
     */
    public void main() throws Exception{
        //Excel文件
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(ResourceUtils.getFile("classpath:台风面雨量.xlsm")));
        //台风面雨量工作表
        XSSFSheet sheet = wb.getSheetAt(0);
        String year = ""; //保存年份

        //循环读取每一条记录，判断其所属年份，计算该年台风总数量并统计影响宁波台风的数量
        for(int i = 1; i <= sheet.getLastRowNum(); i++){
            XSSFRow row = sheet.getRow(i);
            XSSFCell typhooncodeCell = row.getCell(0);
            typhooncodeCell.setCellType(XSSFCell.CELL_TYPE_STRING); //设置单元格格式，否则无法用字符串接收一个将数字格式的cell
            String typhooncode = typhooncodeCell.getStringCellValue();//先取出台风码
            double rainFall = row.getCell(1).getNumericCellValue(); //取出面雨量

            if(!typhooncode.substring(0,4).equals(year)) {  //若当前台风码年份尚未添加进列表中，则新建一个TyphoonInfluence对象
                year = typhooncode.substring(0, 4); //年份为台风码前四位
                TyphoonInfluence typhoonInfluence = new TyphoonInfluence(year,1);
                isRainFallOver50(typhoonInfluence,rainFall,typhooncode);
                typhoonInfluenceList.add(typhoonInfluence);//将该新对象添加到列表中
            }else{ //否则取列表最后一个TyphoonInfluence对象
                TyphoonInfluence typhoonInfluence = typhoonInfluenceList.get(typhoonInfluenceList.size() - 1);
                typhoonInfluence.setNum_of_typhoon(typhoonInfluence.getNum_of_typhoon() + 1); //该年台风数加1
                isRainFallOver50(typhoonInfluence,rainFall,typhooncode);
            }

        }
        calPersentage(); //计算百分比
        writeExel(); //写exel表
    }

    /**
     * 判断面雨量是否大于50
     * @param typhoonInfluence
     * @param rainFall
     * @param typhooncode
     */
    public void isRainFallOver50(TyphoonInfluence typhoonInfluence,double rainFall,String typhooncode){
        if( rainFall > 50){
            typhoonInfluence.setNum_of_nb_typhoon(typhoonInfluence.getNum_of_nb_typhoon() + 1); //如果面雨量大于50，则影响宁波台风数加1
            typhoonInfluence.getTyphooncodes().add(typhooncode); //将该台风码加到台风码列表中
        }
    }

    /**
     * 计算百分比
     */
    public void calPersentage() {
        for (TyphoonInfluence t : typhoonInfluenceList) {  //有的时候forEach循环比for循环更高效，主要看循环中是否需要使用到下标值来决定使用哪一个
            float p = (float) t.getNum_of_nb_typhoon() /  t.getNum_of_typhoon() * 100; // 注意int/int还是int类型，所以要用float/int
            String percentage = String.format("%.1f",p)+"%"; //将float格式化为只保留一位小数并加上百分号
            t.setProportion(percentage);
        }
    }

    /**
     * 将结果写到exel表中
     * @throws Exception
     */
    public void writeExel() throws  Exception{
        String excelFileName = "classppath:影响宁波的台风占比.xlsm";

        String sheetName = "影响宁波的台风";

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        //循环将typhoonInfluenceList的对象写入exel表中（一个对象对应一条记录）
        for (int r = 1; r <= typhoonInfluenceList.size() ; r++)
        {
            TyphoonInfluence typhoonInfluence  = typhoonInfluenceList.get(r - 1);
            XSSFRow row = sheet.createRow(r);

            XSSFCell cell = row.createCell(0);
            cell.setCellValue(typhoonInfluence.getYear());

            XSSFCell cell1 = row.createCell(1);
            cell1.setCellValue(typhoonInfluence.getNum_of_typhoon());

            XSSFCell cell2 = row.createCell(2);
            cell2.setCellValue(typhoonInfluence.getNum_of_nb_typhoon());

            XSSFCell cell3 = row.createCell(3);
            cell3.setCellValue(typhoonInfluence.getProportion());

            XSSFCell cell4 = row.createCell(4);
            ArrayList<String> arr = typhoonInfluence.getTyphooncodes();
            String str = String.join(",",(String[])arr.toArray(new String[arr.size()]));
            cell4.setCellValue(str);
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        // 将workbook写到输出流中
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }
}
