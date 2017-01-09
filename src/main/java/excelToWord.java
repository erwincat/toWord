import freemarker.template.Template;

import java.io.Writer;
import java.util.List;
import java.util.Map;

/**
 * Created by wuchangming on 17/1/6.
 */
public class excelToWord {
    public static  void main(String[] args) throws Exception {
        personnelImportWord();
    }



    private static void personnelImportWord() throws Exception{
        DocumentHandler dh = new DocumentHandler();
        Template t = dh.getTemplate();

        List<Map> list=ReadExcelTable.readTable();
        for (Map map : list){
            Writer out = dh.getWriter((String)map.get("name"));
            dh.createDoc(t,map, out);
        }

    }
}
