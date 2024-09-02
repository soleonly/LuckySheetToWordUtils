package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;

@Data
@ApiModel(value = "LuckySheetDataCellDTO", description = "LuckySheetDataCellDTO实体")
public class LuckySheetDataCTDTO implements Serializable {


    @ApiModelProperty("字体")
    private String fa;

    @ApiModelProperty("")
    private String t;

}
