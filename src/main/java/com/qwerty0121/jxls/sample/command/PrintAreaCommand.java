package com.qwerty0121.jxls.sample.command;

import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

import lombok.Setter;

/**
 * 印刷範囲を設定するコマンド
 */
public class PrintAreaCommand extends AbstractCommand {

  public static final String COMMAND_NAME = "printArea";

  /** 印刷範囲 */
  @Setter
  private String printArea;

  @Override
  public String getName() {
    return COMMAND_NAME;
  }

  @Override
  public Command addArea(Area area) {
    if (areaList.size() >= 1) {
      throw new JxlsException("You can add only a single area to 'printArea' command");
    }
    return super.addArea(area);
  }

  @Override
  public Size applyAt(CellRef cellRef, Context context) {
    if (printArea == null || printArea.isEmpty()) {
      // 印刷範囲を指定していない場合はエラー
      throw new JxlsException("'printArea'を指定してください。");
    }

    // 印刷範囲を設定
    var workbook = ((PoiTransformer) getTransformer()).getWorkbook();
    var sheetIndex = workbook.getSheetIndex(cellRef.getSheetName());
    workbook.setPrintArea(sheetIndex, printArea);

    return areaList.get(0).applyAt(cellRef, context);
  }

}
