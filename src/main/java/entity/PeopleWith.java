package entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@NoArgsConstructor
@AllArgsConstructor
public class PeopleWith {
    @Getter
    @Setter
    private String name;//姓名
    @Getter
    @Setter
    private String Id;//身份证号
    @Getter
    @Setter
    private String familyName;//家庭成员姓名
    @Getter
    @Setter
    private String relation;//称谓
    @Getter
    @Setter
    private
    String familyBirth;//家庭成员出生日期
    @Getter
    @Setter
    private String familyPolitical;//家庭成员政治面貌
    @Getter
    @Setter
    private String familyWorkPlace;//家庭成员工作单位及职务
    @Getter
    @Setter
    private String status;
}
