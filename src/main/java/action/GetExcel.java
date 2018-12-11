package action;

import entity.People;
import entity.PeopleWith;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class GetExcel {
    public Map<String, List<People>> readExcel() {
        Map<String, List<People>> map = new HashMap<>();
        List<People> list = new ArrayList<>();
        List<People> peopleList = new ArrayList<>();
        List<PeopleWith> peopleWithList = new ArrayList<>();
        List<People> peopleList1 = new ArrayList<>();
        List<People> peopleList2 = new ArrayList<>();
        List<People> peopleList3 = new ArrayList<>();
        File file = new File("D:\\Excel信息导入导出.xls");
        String[] temp;
        // 设置编码格式
        WorkbookSettings setEncode = new WorkbookSettings();
        setEncode.setEncoding("UTF-8");
        try {
            Workbook wb = Workbook.getWorkbook(file, setEncode);
            // 0代表第一个工作表对象
            Sheet sheet1 = wb.getSheet(2);
            Sheet sheet2 = wb.getSheet(9);
            //工作表1的行数
            int rows1 = sheet1.getRows();
            //工作表2的行数
            int rows2 = sheet2.getRows();
            for (int i = 2; i < rows1; i++) {
                if (sheet1.getCell(0, i) == null) {
                    break;
                }
                People people = new People();
                people.setName(sheet1.getCell(0, i).getContents());
                people.setId(sheet1.getCell(1, i).getContents());
                people.setSex(sheet1.getCell(2, i).getContents());
                people.setBirthday(sheet1.getCell(3, i).getContents());
                people.setJobInfo(sheet1.getCell(19, i).getContents());
                people.setJobLocation(sheet1.getCell(27, i).getContents());
                list.add(people);
            }
            for (int i = 0; i < list.size(); i++) {
                for (int j = i + 1; j < list.size(); j++) {
                    if (list.get(i).getName().equals(list.get(j).getName())
                        && list.get(i).getId().equals(list.get(j).getId())) {
                        list.remove(i);
                    }
                }
            }
            for (int j = 2; j < rows2; j++) {
                PeopleWith peopleWith = new PeopleWith();
                if ((
                        ("妻子").equals(sheet2.getCell(3, j).getContents())
                        || ("丈夫").equals(sheet2.getCell(3, j).getContents()) ||
                        ("配偶").equals(sheet2.getCell(3, j).getContents())
                )) {
                    if (sheet2.getCell(6, j).getContents().contains("陵川")) {
                        peopleWith.setStatus("陵川");
                    } else if (
                            sheet2.getCell(6, j).getContents().contains(
                                    "高平")
                    ) {
                        peopleWith.setStatus("高平");
                    } else if (
                            sheet2.getCell(6, j).getContents().contains(
                                    "沁水")
                    ) {
                        peopleWith.setStatus("沁水");
                    } else if (sheet2.getCell(6, j).getContents().contains(
                            "阳城")) {
                        peopleWith.setStatus("阳城");
                    } else if (sheet2.getCell(6, j).getContents().contains(
                            "泽州")) {
                        peopleWith.setStatus("泽州");
                    } else if
                    ((
                             sheet2.getCell(6, j).getContents().contains("晋城")
                             || sheet2.getCell(6, j).getContents().contains("市")
                     )
                     && !sheet2.getCell(6, j).getContents().contains("陵川")
                     && !sheet2.getCell(6, j).getContents().contains("高平")
                     && !sheet2.getCell(6, j).getContents().contains("沁水")
                     && !sheet2.getCell(6, j).getContents().contains("泽州")
                     && !sheet2.getCell(6, j).getContents().contains("阳城")) {
                        peopleWith.setStatus("晋城");
                    } else {
                        continue;
                    }
                    peopleWith.setName(sheet2.getCell(0, j).getContents());
                    peopleWith.setId(sheet2.getCell(1, j).getContents());
                    peopleWith.setFamilyName(sheet2.getCell(2, j).getContents());
                    peopleWith.setRelation(sheet2.getCell(3, j).getContents());
                    peopleWith.setFamilyBirth(sheet2.getCell(4, j).getContents());
                    peopleWith.setFamilyPolitical(sheet2.getCell(5, j).getContents());
                    peopleWith.setFamilyWorkPlace(sheet2.getCell(6, j).getContents());
                    peopleWithList.add(peopleWith);
                }
            }
            for (int i = 0; i < list.size(); i++) {
                for (int j = 0; j < peopleWithList.size(); j++) {
                    if ((peopleWithList.get(j).getId().equals(list.get(i).getId()))
                        && (peopleWithList.get(j).getName().equals(list.get(i).getName()))) {
                        temp = list.get(i).getJobLocation().split(" ");
                        if (temp[1] != null) {
                            if (temp[1].endsWith("县")) {
                                if (temp[1].contains("晋城")
                                    && "晋城".equals(peopleWithList.get(j)
                                        .getStatus())) {
                                    list.get(i).setPeopleWith(peopleWithList.get(j));
                                    peopleList.add(list.get(i));
                                    peopleList1.add(list.get(i));
                                } else if (!temp[1].contains(peopleWithList.get(j).getStatus())) {
                                    list.get(i).setPeopleWith(peopleWithList.get(j));
                                    peopleList.add(list.get(i));
                                    peopleList1.add(list.get(i));
                                }
                            } else if (temp[1].endsWith("高平市")) {
                                if (temp[1].contains("晋城")
                                    && "晋城".equals(peopleWithList.get(j)
                                        .getStatus())) {
                                    list.get(i).setPeopleWith(peopleWithList.get(j));
                                    peopleList.add(list.get(i));
                                    peopleList1.add(list.get(i));
                                } else if (!"高平".equals(peopleWithList.get(j).getStatus())) {
                                    list.get(i).setPeopleWith(peopleWithList.get(j));
                                    peopleList.add(list.get(i));
                                    peopleList1.add(list.get(i));
                                }
                            } else {
                                if (!("晋城".equals(peopleWithList.get(j).getStatus()))) {
                                    list.get(i).setPeopleWith(peopleWithList.get(j));
                                    peopleList.add(list.get(i));
                                    peopleList1.add(list.get(i));
                                }
                            }
                        }
                    }
                }
            }
            //peopleList晋城市所有异地夫妻的集合 ,peopleWithList所有在晋城市的夫妻集合
            System.out.println(peopleList.get(3).getName());
            //peopleList内的peopleWith为配偶关系对象，peopleList为配偶关系集
            //familyName为配偶姓名，Name为公务员姓名
            for (int n = 0; n < peopleList.size(); n++) {
                for (int j = n + 1; j < peopleList.size(); j++) {
                    if (peopleList.get(n).getPeopleWith().getFamilyName().equals(peopleList.get(j).getName())
                        && peopleList.get(j).getPeopleWith().getFamilyName().equals(peopleList.get(n).getName())) {
                        peopleList2.add(peopleList.get(n));
                        peopleList3.add(peopleList.get(j));
                    }
                }
            }
            peopleList1.removeAll(peopleList2);
            peopleList1.removeAll(peopleList3);
            //peopleList.removeAll(peopleList2);
            map.put("peopleList", peopleList);//所有的异地集合
            map.put("peopleList1", peopleList1);//一个是公务员 一个不是
            map.put("peopleList2", peopleList2);//都是公务员
        } catch (Exception e) {
            e.printStackTrace();
        }
        People people = list.get(6643);
        return map;
    }
}

