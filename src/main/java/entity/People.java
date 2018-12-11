package entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@AllArgsConstructor
@NoArgsConstructor
public class People {
    @Getter
    @Setter
    private String name;//姓名
    @Getter
    @Setter
    private String Id;//身份证号
    @Getter
    @Setter
    private String sex;//性别
    @Getter
    @Setter
    private String birthday;//出生年月
    @Getter
    @Setter
    private String jobInfo;//现工作单位及职务
    @Getter
    @Setter
    private String jobLocation;//所在单位
    @Getter
    @Setter
    private PeopleWith peopleWith;
}
