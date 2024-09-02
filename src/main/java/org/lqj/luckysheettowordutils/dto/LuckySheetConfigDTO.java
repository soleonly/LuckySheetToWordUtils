package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

@Data
@ApiModel(value = "LuckySheetConfigDTO", description = "LuckySheetConfigDTO实体")
public class LuckySheetConfigDTO implements Serializable {

    @ApiModelProperty("边框样式")
    private List<LuckySheetBorderInfoDTO> borderInfo;
    @ApiModelProperty("合并信息")
    private Map<String,LuckySheetMergeDTO> merge;
}
