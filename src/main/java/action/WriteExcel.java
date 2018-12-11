package action;

import entity.People;
import jxl.Workbook;
import jxl.format.*;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class WriteExcel {
    public void writeExcel(Map<String, List<People>> map) throws IOException,
            WriteException {
        WritableFont wc_title =
                new WritableFont(WritableFont.createFont("Arial"), 18,
                        WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE,
                        Colour.BLACK);
        WritableCellFormat wcf_title = new WritableCellFormat(wc_title);
        wcf_title.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
        wcf_title.setAlignment(Alignment.CENTRE);
        wcf_title.setVerticalAlignment(VerticalAlignment.CENTRE);
        wcf_title.setWrap(true);
        WritableFont wc_name = new WritableFont(WritableFont.createFont("黑体")
                , 10, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE,
                Colour.RED);
        WritableCellFormat wcf_name = new WritableCellFormat(wc_name);
        wcf_name.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
        wcf_name.setAlignment(Alignment.CENTRE);
        wcf_name.setVerticalAlignment(VerticalAlignment.CENTRE);
        wcf_name.setWrap(true);
        WritableFont wc_body = new WritableFont(WritableFont.createFont("仿宋_GB2312")
                , 10, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE,
                Colour.BLACK);
        WritableCellFormat wcf_body = new WritableCellFormat(wc_body);
        wcf_body.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLACK);
        wcf_body.setAlignment(Alignment.CENTRE);
        wcf_body.setVerticalAlignment(VerticalAlignment.CENTRE);
        wcf_body.setWrap(true);
        List<People> peopleList = map.get("peopleList");
        List<People> peopleList1 = map.get("peopleList1");
        List<People> peopleList2 = map.get("peopleList2");
        String path = "D://Excel统计结果.xls";
        Label label = null;
        File file = new File(path);
        WritableWorkbook workbook = Workbook.createWorkbook(file);// 找到源文件进行类型转换
        WritableSheet sheet = workbook.createSheet("干部信息表", 0);
        // 设置titles 第一行
        String title = "晋城市配偶异地工作干部信息表";
        label = new Label(0, 0, title, wcf_title);
        sheet.addCell(label);
        sheet.mergeCells(0, 0, 10, 0);
        // 设置第二行
        String titles[] = {"姓名", "身份证号", "性别", "出生年月", "现工作单位及职务全称", "统计关系所在单位", "家庭成员姓名", "称谓", "出生日期", "政治面貌", "工作单位及职务"};
        for (int i = 0; i < titles.length; i++) {
            label = new Label(i, 1, titles[i], wcf_name);
            sheet.addCell(label);
        }
        //遍历集合 填充数据
        for (int i = 0; i < peopleList.size(); i++) {
            Label label1 = new Label(0,
                    i + 2, peopleList.get(i).getName(), wcf_body);
            Label label2 = new Label(1,
                    i + 2, peopleList.get(i).getId(), wcf_body);
            Label label3 = new Label(2,
                    i + 2, peopleList.get(i).getSex(), wcf_body);
            Label label4 = new Label(3, i + 2,
                    peopleList.get(i).getBirthday(), wcf_body);
            Label label5 = new Label(4,
                    i + 2, peopleList.get(i).getJobInfo(), wcf_body);
            Label label6 = new Label(5, i + 2,
                    peopleList.get(i).getJobLocation(), wcf_body);
            Label label7 = new Label(6,
                    i
                    + 2, peopleList.get(i).getPeopleWith().getFamilyName(), wcf_body);
            Label label8 = new Label(7,
                    i
                    + 2, peopleList.get(i).getPeopleWith().getRelation(), wcf_body);
            Label label9 = new Label(8,
                    i
                    + 2, peopleList.get(i).getPeopleWith().getFamilyBirth(), wcf_body);
            Label label10 = new Label(9, i
                                         + 2, peopleList.get(i).getPeopleWith().getFamilyPolitical(), wcf_body);
            Label label11 = new Label(10, i
                                          + 2, peopleList.get(i).getPeopleWith().getFamilyWorkPlace(), wcf_body);
            sheet.addCell(label1);
            sheet.addCell(label2);
            sheet.addCell(label3);
            sheet.addCell(label4);
            sheet.addCell(label5);
            sheet.addCell(label6);
            sheet.addCell(label7);
            sheet.addCell(label8);
            sheet.addCell(label9);
            sheet.addCell(label10);
            sheet.addCell(label11);
        }
        //行高
        sheet.setRowView(0, 700);
        sheet.setRowView(1, 550);
        //列宽
        sheet.setColumnView(0, 10);
        sheet.setColumnView(1, 24);
        sheet.setColumnView(2, 7);
        sheet.setColumnView(3, 10);
        sheet.setColumnView(4, 52);
        sheet.setColumnView(5, 52);
        sheet.setColumnView(6, 13);
        sheet.setColumnView(7, 14);
        sheet.setColumnView(8, 14);
        sheet.setColumnView(9, 14);
        sheet.setColumnView(10, 38);
        WritableSheet sheet1 = workbook.createSheet("干部信息表(配偶非干部)", 1);
        // 设置titles 第一行
        String title1 = "晋城市配偶异地工作干部信息表(配偶非干部)";
        label = new Label(0, 0, title, wcf_title);
        sheet1.addCell(label);
        sheet1.mergeCells(0, 0, 10, 0);
        // 设置第二行
        String titles1[] = {"姓名", "身份证号", "性别", "出生年月", "现工作单位及职务全称",
                "统计关系所在单位", "家庭成员姓名", "称谓", "出生日期", "政治面貌", "工作单位及职务"};
        for (int i = 0; i < titles1.length; i++) {
            label = new Label(i, 1, titles1[i],wcf_name);
            sheet1.addCell(label);
        }
        //遍历集合 填充数据
        for (int i = 0; i < peopleList1.size(); i++) {
            Label label1 = new Label(0,
                    i + 2, peopleList1.get(i).getName(), wcf_body);
            Label label2 = new Label(1,
                    i + 2, peopleList1.get(i).getId(), wcf_body);
            Label label3 = new Label(2,
                    i + 2, peopleList1.get(i).getSex(), wcf_body);
            Label label4 = new Label(3, i + 2,
                    peopleList1.get(i).getBirthday(), wcf_body);
            Label label5 = new Label(4,
                    i + 2, peopleList1.get(i).getJobInfo(), wcf_body);
            Label label6 = new Label(5, i + 2,
                    peopleList1.get(i).getJobLocation(), wcf_body);
            Label label7 = new Label(6,
                    i
                    + 2, peopleList1.get(i).getPeopleWith().getFamilyName(), wcf_body);
            Label label8 = new Label(7,
                    i
                    + 2, peopleList1.get(i).getPeopleWith().getRelation(), wcf_body);
            Label label9 = new Label(8,
                    i
                    + 2, peopleList1.get(i).getPeopleWith().getFamilyBirth(), wcf_body);
            Label label10 = new Label(9, i
                                         + 2,
                    peopleList1.get(i).getPeopleWith().getFamilyPolitical(), wcf_body);
            Label label11 = new Label(10, i
                                          + 2,
                    peopleList1.get(i).getPeopleWith().getFamilyWorkPlace(), wcf_body);
            sheet1.addCell(label1);
            sheet1.addCell(label2);
            sheet1.addCell(label3);
            sheet1.addCell(label4);
            sheet1.addCell(label5);
            sheet1.addCell(label6);
            sheet1.addCell(label7);
            sheet1.addCell(label8);
            sheet1.addCell(label9);
            sheet1.addCell(label10);
            sheet1.addCell(label11);
        }
        //行高
        sheet1.setRowView(0, 700);
        sheet1.setRowView(1, 550);
        //列宽
        sheet1.setColumnView(0, 10);
        sheet1.setColumnView(1, 24);
        sheet1.setColumnView(2, 7);
        sheet1.setColumnView(3, 10);
        sheet1.setColumnView(4, 52);
        sheet1.setColumnView(5, 52);
        sheet1.setColumnView(6, 13);
        sheet1.setColumnView(7, 14);
        sheet1.setColumnView(8, 14);
        sheet1.setColumnView(9, 14);
        sheet1.setColumnView(10, 38);
        WritableSheet sheet2 = workbook.createSheet("干部信息表(配偶是干部)", 2);
        // 设置titles 第一行
        String title2 = "晋城市配偶异地工作干部信息表(配偶是干部)";
        label = new Label(0, 0, title2, wcf_title);
        sheet2.addCell(label);
        sheet2.mergeCells(0, 0, 10, 0);
        // 设置第二行
        String titles2[] = {"姓名", "身份证号", "性别", "出生年月", "现工作单位及职务全称",
                "统计关系所在单位", "家庭成员姓名", "称谓", "出生日期", "政治面貌", "工作单位及职务"};
        for (int i = 0; i < titles2.length; i++) {
            label = new Label(i, 1, titles2[i], wcf_name);
            sheet2.addCell(label);
        }
        //遍历集合 填充数据
        for (int i = 0; i < peopleList2.size(); i++) {
            Label label1 = new Label(0,
                    i + 2, peopleList2.get(i).getName(), wcf_body);
            Label label2 = new Label(1,
                    i + 2, peopleList2.get(i).getId(), wcf_body);
            Label label3 = new Label(2,
                    i + 2, peopleList2.get(i).getSex(), wcf_body);
            Label label4 = new Label(3, i + 2,
                    peopleList2.get(i).getBirthday(), wcf_body);
            Label label5 = new Label(4,
                    i + 2, peopleList2.get(i).getJobInfo(), wcf_body);
            Label label6 = new Label(5, i + 2,
                    peopleList2.get(i).getJobLocation(), wcf_body);
            Label label7 = new Label(6,
                    i
                    + 2, peopleList2.get(i).getPeopleWith().getFamilyName(), wcf_body);
            Label label8 = new Label(7,
                    i
                    + 2, peopleList2.get(i).getPeopleWith().getRelation(), wcf_body);
            Label label9 = new Label(8,
                    i
                    + 2, peopleList2.get(i).getPeopleWith().getFamilyBirth(), wcf_body);
            Label label10 = new Label(9, i
                                         + 2,
                    peopleList2.get(i).getPeopleWith().getFamilyPolitical(), wcf_body);
            Label label11 = new Label(10, i
                                          + 2,
                    peopleList2.get(i).getPeopleWith().getFamilyWorkPlace(), wcf_body);
            sheet2.addCell(label1);
            sheet2.addCell(label2);
            sheet2.addCell(label3);
            sheet2.addCell(label4);
            sheet2.addCell(label5);
            sheet2.addCell(label6);
            sheet2.addCell(label7);
            sheet2.addCell(label8);
            sheet2.addCell(label9);
            sheet2.addCell(label10);
            sheet2.addCell(label11);
        }
        //行高
        sheet2.setRowView(0, 700);
        sheet2.setRowView(1, 550);
        //列宽
        sheet2.setColumnView(0, 10);
        sheet2.setColumnView(1, 24);
        sheet2.setColumnView(2, 7);
        sheet2.setColumnView(3, 10);
        sheet2.setColumnView(4, 52);
        sheet2.setColumnView(5, 52);
        sheet2.setColumnView(6, 13);
        sheet2.setColumnView(7, 14);
        sheet2.setColumnView(8, 14);
        sheet2.setColumnView(9, 14);
        sheet2.setColumnView(10, 38);
        workbook.write(); // 保存文件
        workbook.close(); // 关闭文件
    }
}
