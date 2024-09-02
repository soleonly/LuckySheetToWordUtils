package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;

@Data
@ApiModel(value = "LuckySheetBorderInfoValueDTO", description = "LuckySheetBorderInfoValueDTO实体")
public class LuckySheetBorderInfoValueDTO implements Serializable {

    @ApiModelProperty("列号")
    private int col_index;
    @ApiModelProperty("行号")
    private int row_index;

    private LuckySheetBorderInfoValueItemDTO b;
    private LuckySheetBorderInfoValueItemDTO l;
    private LuckySheetBorderInfoValueItemDTO r;
    private LuckySheetBorderInfoValueItemDTO t;
}


