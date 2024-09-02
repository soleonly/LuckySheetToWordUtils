package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;
import java.util.List;

@Data
@ApiModel(value = "LuckySheetDataDTO", description = "LuckySheetDataDTO实体")
public class LuckySheetDTO implements Serializable {
    @ApiModelProperty("数据")
    private List<List<LuckySheetDataCellDTO>> data;
    @ApiModelProperty("配置")
    private LuckySheetConfigDTO config;
}
