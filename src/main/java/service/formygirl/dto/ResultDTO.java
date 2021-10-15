package service.formygirl.dto;

import lombok.Data;

/**
 * @author kq644
 * @since 2021-10-14 17:25
 */
@Data
public class ResultDTO {
    String name;
    Integer grade;
    String progress;
    String code;
    String phone;

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
}
