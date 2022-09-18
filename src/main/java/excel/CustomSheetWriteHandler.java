package excel;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * @author by liangzj
 * @since 2022/9/17 16:08
 */
public class CustomSheetWriteHandler implements SheetWriteHandler {

    @Override
    public void afterSheetCreate(
            WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        // 设置保护密码
        writeSheetHolder.getSheet().protectSheet("123456");
        // 不允许表格复制，防止别人复制表格到别的excel中修改
        ((SXSSFSheet) writeSheetHolder.getSheet()).lockSelectLockedCells(true);
    }
}
