package data;

import lombok.Data;

@Data
public class MapDTO {
    int orderName;
    int orderUnit;
    int orderStudentNum;
    int orderStudentUnitPrice;
    int orderStudentTotalPrice;
    int orderTeacherNum;
    int orderTeacherUnitPrice;
    int orderTeacherTotalPrice;
    boolean hasTeacherInfo;

    public void setNumber(int i, int count) {
        if (count == 1) {
            this.orderUnit = i - 1;
            this.orderName = i - 2;
        } else if (count == 2) {
            this.orderStudentNum = i;
        } else if (count == 3) {
            this.orderTeacherNum = i;
        }
    }

    public void setUnitPrice(int i, int count) {
        if (count == 1) {
        } else if (count == 2) {
            this.orderStudentUnitPrice = i;
        } else if (count == 3) {
            this.orderTeacherTotalPrice = i;
        }
    }

    public void setTotalPrice(int i, int count) {
        if (count == 1) {
        } else if (count == 2) {
            this.orderStudentTotalPrice = i;
        } else if (count == 3) {
            this.orderTeacherTotalPrice = i;
        }
    }
}
