package service.formygirl.dto;

import java.util.Objects;
import lombok.Data;

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

    public String getDepart() {
        return depart == null ? "未知部门" : depart;
    }

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

    public boolean pass() {
        if (Objects.isNull(progress)) {
            return false;
        }
        double progressD = Double.parseDouble(this.progress.replace("%", ""));
        if (grade < 80 || progressD < 90) {
            return false;
        }
        return true;
    }

    public String findCheckMessage() {
        return pass() ? "通过" : "不通过";
    }
}
