package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;

@Data
@ApiModel(value = "LuckySheetDataCellDTO", description = "LuckySheetDataCellDTO实体")
public class LuckySheetDataCellDTO implements Serializable {
    @ApiModelProperty("")
    private Integer bl;

    @ApiModelProperty("字体大小")
    private Integer fs;

    @ApiModelProperty("")
    private Integer sjbs;

    @ApiModelProperty("")
    private String ht;

    @ApiModelProperty("标记")
    private String m;

    @ApiModelProperty("值")
    private String v;

    @ApiModelProperty("合并信息")
    private LuckySheetMergeDTO mc;

    @ApiModelProperty("字体信息")
    private LuckySheetDataCTDTO ct;

}
