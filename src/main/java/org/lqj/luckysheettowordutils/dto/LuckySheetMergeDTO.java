package org.lqj.luckysheettowordutils.dto;

import io.swagger.annotations.ApiModel;
import lombok.Data;

import java.io.Serializable;

@Data
@ApiModel(value = "LuckySheetMergeDTO", description = "LuckySheetMergeDTO实体")
public class LuckySheetMergeDTO implements Serializable {

    private int c;
    private int r;
    private int cs;
    private int rs;

}
