package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import lombok.Data;

import java.io.Serializable;

@Data
@ApiModel(value = "LuckySheetBorderInfoValueItemDTO", description = "LuckySheetBorderInfoValueItemDTO实体")
public class LuckySheetBorderInfoValueItemDTO implements Serializable {
    private String corlor;
    private String style;

}


