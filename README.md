# police-resource : 경찰청 인적 자원 정보를 전처리 프로그램
1. 원본 데이터는 기밀 사항이므로 공개하지 않는다.
2. 경찰청에 근무 중인 약 1500명 정도의 종사자들의 개인정보와 교육 이력, 전문수사관/마스터 취득 이력, 자격증, 학력 등을 정제하여 저장한다.
3. 개인 정보를 제외한 모든 데이터는 입직 전 ~ 30년차까지 연차 별로 구분되어 저장한다.
4. 교육 이력은 키워드 포함 여부에 따라 기초과정, 기본과정, 전문 과정으로 구분하여 정수 인코딩된 형태로 저장한다.
5. 전문 수사관, 마스터 취득 이력은 그 종류에 따라 정수 인코딩하여 저장한다.
6. 학력은 없음, 학사, 석사, 박사에 따라 정수 인코딩하여 저장한다.
7. 연구 실적은 연구실적/공모전 콜럼에서 분야가 일치하는 것 만을 고려한다.
8. 나머지 데이터는 종류에 상관 없이 그 개수만을 고려한다.
9. 여러 곳에서 근무한 경력이 있을 경우, 어떤 근무지의 근무 기간과 다른 근무지의 근무 기간이 겹치는 경우 해당 행을 하이라이팅 처리한다.
10. 한 사람의 데이터가 여러 줄에 걸쳐서 저장되어 있으므로, 데이터를 한 줄 씩 읽으면서 데이터의 인적 사항이 바뀌는 경우를 고려한다.
11. 추가적으로 굵기, 셀 안에서의 위치 정렬을 고려한다.
