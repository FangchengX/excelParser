package service.formygirl.dto;

import lombok.Data;

import java.util.Objects;

/**
 * @author kq644
 * @since 2021-10-14 17:25
 */
@Data
public class ResultDTO {
    String name;
    double grade;
    String progress;
    String code;
    String phone;
    String account;
    String depart;
    String id;

    public void setGrade(double grade) {
        this.grade = Math.max(grade, this.grade);
    }

    public void setProgress(String progress) {
        if (Objects.isNull(this.progress)) {
            this.progress = progress;
        } else {
            double progressD = Double.parseDouble(this.progress.replace("%", ""));
            double progress1 = Double.parseDouble(progress.replace("%", ""));
            if (progress1 > progressD) {
                this.progress = progress;
            }
        }
    }

    public void setDepart(String depart) {
        if (Objects.isNull(this.depart)) {
            this.depart = depart;
        }
    }

    @Override
    public String toString() {
        return "ResultDTO{" +
            "name='" + name + '\'' +
            ", grade=" + grade +
            ", progress='" + progress + '\'' +
            ", code='" + code + '\'' +
            ", phone='" + phone + '\'' +
            '}';
    }

    public String findCheckMessage() {
        if (Objects.isNull(progress)) {
            return "不通过";
        }
        double progressD = Double.parseDouble(this.progress.replace("%", ""));
        if (grade < 80 || progressD < 90) {
            return "不通过";
        }
        return "通过";
    }
}
