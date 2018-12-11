package action;

import entity.People;
import jxl.write.WriteException;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class run {
    public static void main(String[] args) throws IOException, WriteException {
        GetExcel getExcel = new GetExcel();
        Map<String, List<People>> map= getExcel.readExcel();
        WriteExcel writeExcel=new WriteExcel();
        writeExcel.writeExcel(map);
    }
}
