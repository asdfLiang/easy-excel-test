package excel;

import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 单元格样式处理器
 *
 * @author by liangzj
 * @since 2022/9/17 16:26
 */
public class StyleWriteHandler extends LongestMatchColumnWidthStyleStrategy {

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        if (context.getHead()) {
            headerStyle(context);
        } else {
            //            contentStyle(context);
            contentStyle2(context.getCell());
        }
    }

    /**
     * 表数据格式处理1
     *
     * @param context
     */
    private void contentStyle(CellWriteHandlerContext context) {
        // 解锁没有内容的单元格(方法1)
        WriteCellStyle writeCellStyle = context.getFirstCellData().getOrCreateStyle();
        /* !! 注意：这行就是解锁单元格的代码，locked == true为锁定，locked == false为不锁定 */
        writeCellStyle.setLocked(StringUtils.isNotBlank(context.getCell().getStringCellValue()));
        // 如果锁定，置灰
        if (writeCellStyle.getLocked()) {
            WriteFont writeFont = new WriteFont();
            writeFont.setColor(IndexedColors.GREY_40_PERCENT.index);
            writeCellStyle.setWriteFont(writeFont);
        }
    }

    /**
     * 表数据格式处理2
     *
     * @param cell
     */
    private void contentStyle2(Cell cell) {
        // 创建新的单元格样式(方法2)
        CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        // 复制原来单元格的样式
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        /* !! 注意：这行就是解锁单元格的代码，locked == true为锁定，locked == false为不锁定 */
        cellStyle.setLocked(StringUtils.isNotBlank(cell.getStringCellValue()));
        cell.setCellStyle(cellStyle);
        // 如果锁定，置灰
        if (cell.getCellStyle().getLocked()) {
            Font font = cell.getSheet().getWorkbook().createFont();
            font.setColor(IndexedColors.GREY_40_PERCENT.index);
            cellStyle.setFont(font);
        }
    }

    /**
     * 表头格式处理
     *
     * @param context
     */
    private static void headerStyle(CellWriteHandlerContext context) {
        Cell cell = context.getCell();

        int colWidth = cell.getStringCellValue().length() * 1500;
        boolean needHidden = "requireId".equals(cell.getStringCellValue());

        // 根据表头文字设置列宽
        cell.getSheet().setColumnWidth(cell.getColumnIndex(), colWidth);
        // 冻结表头
        cell.getSheet().createFreezePane(1, 2);
        // 隐藏指定列
        cell.getSheet().setColumnHidden(cell.getColumnIndex(), needHidden);
    }
}
