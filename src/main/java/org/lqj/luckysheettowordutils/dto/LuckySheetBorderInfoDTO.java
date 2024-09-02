package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;

@Data
@ApiModel(value = "LuckySheetBorderInfoDTO", description = "LuckySheetBorderInfoDTO实体")
public class LuckySheetBorderInfoDTO implements Serializable {

    @ApiModelProperty("边框类型")
    private String borderType;
    @ApiModelProperty("范围类型")
    private String rangeType;
    @ApiModelProperty("颜色")
    private String corlor;
    @ApiModelProperty("样式")
    private String style;
    @ApiModelProperty("值")
    private LuckySheetBorderInfoValueDTO value;
}


