package service.formygirl.dto;

import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author kq644
 * @since 2021-10-28 11:09
 */
@Data
@NoArgsConstructor
public class DepartDTO {
    String depart;
    int passNumber;
    int failNumber;
    int totalNumber;
    double passPercent;

    public DepartDTO(String depart) {
        this.depart = depart;
        this.passNumber = 0;
        this.failNumber = 0;
        this.totalNumber = 0;
        this.passPercent = 0;
    }

    public void addNumber(boolean pass) {
        if (pass) {
            passNumber++;
        } else {
            failNumber++;
        }
        totalNumber++;
        this.passPercent = 1.0 * passNumber / totalNumber;
    }
}
