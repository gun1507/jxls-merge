package com.jxls.jxlsmerge;

import com.jxls.jxlsmerge.domain.Group;
import com.jxls.jxlsmerge.domain.Name;
import com.jxls.jxlsmerge.domain.Pass;
import com.jxls.jxlsmerge.domain.School;
import com.jxls.jxlsmerge.excel.ExcelGenerateException;
import com.jxls.jxlsmerge.excel.ExcelUtil;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

@SpringBootTest
public class JxlsMergeTests {

    private final ExcelUtil excelUtil;

    private List<School> schoolList;
    private List<Pass> passList;

    public JxlsMergeTests() {
        excelUtil = new ExcelUtil();
        schoolList = null;
        passList = new ArrayList<>();
    }

    @BeforeEach
    void setUp() {
        schoolList = List.of(
                School.builder()
                        .name("일반고")
                        .groupList(List.of(
                                Group.builder()
                                        .name("일반")
                                        .nameList(List.of(
                                                Name.builder().name("김가영").score(43).build(),
                                                Name.builder().name("이나은").score(60).build(),
                                                Name.builder().name("박다인").score(80).build(),
                                                Name.builder().name("최라은").score(100).build(),
                                                Name.builder().name("정마은").score(93).build()
                                        ))
                                        .build()
                        ))
                        .build(),
                School.builder()
                        .name("특목고")
                        .groupList(List.of(
                                Group.builder()
                                        .name("과학")
                                        .nameList(List.of(
                                                Name.builder().name("강바영").score(54).build(),
                                                Name.builder().name("윤사은").score(65).build(),
                                                Name.builder().name("장아린").score(98).build(),
                                                Name.builder().name("한자윤").score(69).build(),
                                                Name.builder().name("배찬영").score(80).build()
                                        ))
                                        .build(),
                                Group.builder()
                                        .name("외국어")
                                        .nameList(List.of(
                                                Name.builder().name("오태현").score(79).build(),
                                                Name.builder().name("김파은").score(88).build(),
                                                Name.builder().name("이하린").score(89).build(),
                                                Name.builder().name("박건우").score(38).build(),
                                                Name.builder().name("최경민").score(78).build()
                                        ))
                                        .build(),
                                Group.builder()
                                        .name("국제")
                                        .nameList(List.of(
                                                Name.builder().name("김두리").score(77).build(),
                                                Name.builder().name("이미주").score(78).build(),
                                                Name.builder().name("박바다").score(99).build(),
                                                Name.builder().name("장서연").score(18).build(),
                                                Name.builder().name("김수아").score(98).build()
                                        ))
                                        .build(),
                                Group.builder()
                                        .name("예술")
                                        .nameList(List.of(
                                                Name.builder().name("이시현").score(32).build(),
                                                Name.builder().name("정아름").score(56).build(),
                                                Name.builder().name("강용준").score(66).build(),
                                                Name.builder().name("박유진").score(100).build(),
                                                Name.builder().name("김이랑").score(99).build()
                                        ))
                                        .build(),
                                Group.builder()
                                        .name("체육")
                                        .nameList(List.of(
                                                Name.builder().name("정재민").score(70).build(),
                                                Name.builder().name("이정호").score(67).build(),
                                                Name.builder().name("한지민").score(44).build(),
                                                Name.builder().name("오찬미").score(45).build(),
                                                Name.builder().name("김태훈").score(78).build()
                                        ))
                                        .build(),
                                Group.builder()
                                        .name("마이스터")
                                        .nameList(List.of(
                                                Name.builder().name("최하린").score(88).build(),
                                                Name.builder().name("정승우").score(89).build(),
                                                Name.builder().name("강민우").score(19).build(),
                                                Name.builder().name("이예진").score(100).build(),
                                                Name.builder().name("박승민").score(88).build()
                                        ))
                                        .build()
                        ))
                        .build(),
                School.builder()
                        .name("특성화고")
                        .groupList(List.of(
                                Group.builder()
                                        .name("특성")
                                        .nameList(List.of(
                                                Name.builder().name("김윤호").score(77).build(),
                                                Name.builder().name("장주연").score(70).build(),
                                                Name.builder().name("한세진").score(80).build(),
                                                Name.builder().name("오은희").score(70).build(),
                                                Name.builder().name("윤민수").score(89).build()
                                        ))
                                        .build()
                        ))
                        .build(),
                School.builder()
                        .name("자율고")
                        .groupList(List.of(
                                Group.builder()
                                        .name("자율")
                                        .nameList(List.of(
                                                Name.builder().name("김지현").score(89).build(),
                                                Name.builder().name("이영재").score(99).build(),
                                                Name.builder().name("박민주").score(19).build(),
                                                Name.builder().name("최은영").score(90).build(),
                                                Name.builder().name("정태준").score(97).build()
                                        ))
                                        .build()
                        ))
                        .build()
        );
    }

    @Test
    void merge_test() {
        try {
            for(School school : schoolList) {
                for(Group group : school.getGroupList()) {
                    for(Name name : group.getNameList()) {
                        if(name.getScore() > 70) {
                            passList.add(Pass.builder().name("합격").build());
                        }else{
                            passList.add(Pass.builder().name("불합격").build());
                        }
                    }
                }
            }
            excelUtil.writeExcelFile("result.xlsx", schoolList, passList);
        } catch (Exception e) {
            throw new ExcelGenerateException("엑셀 파일 생성 실패 하였습니다.", e);
        }
    }

}
