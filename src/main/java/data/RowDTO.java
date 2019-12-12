package data;

import lombok.Data;

/**
 * Created by kq644 on 2019/12/9.
 */
@Data
public class RowDTO {
    String sheetName;
    Integer no;
    String name;
    String unit;
    Integer number;
    Double unitPrice;
    Double totalPrice;
    String comment;
}
