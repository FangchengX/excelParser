package service.exam;

import lombok.Data;

/**
 * @author kq644
 * @since 2020-11-02 13:07
 */
@Data
public class StudentDTO {
    int number;
    String name;
    String clazz;
    /**
     * 学号
     */
    String code;
}
