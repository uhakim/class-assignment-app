# -*- coding: utf-8 -*-
"""
반편성 배정 프로그램 - Streamlit 웹 앱
"""
import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
import sys
import os
import shutil

# 기존 모듈 import는 직접 실행 방식으로 변경
import subprocess
import tempfile
import shutil

# exe로 패키징된 경우 경로 처리
if getattr(sys, 'frozen', False):
    # exe로 실행 중인 경우
    application_path = os.path.dirname(sys.executable)
    os.chdir(application_path)
else:
    # 스크립트로 실행 중인 경우
    application_path = os.path.dirname(os.path.abspath(__file__))
    os.chdir(application_path)

sys.stdout.reconfigure(encoding='utf-8')

st.set_page_config(
    page_title="반편성 배정 프로그램",
    page_icon="📚",
    layout="wide"
)

st.title("📚 반편성 배정 프로그램")
st.markdown("---")

# 세션 상태 초기화
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'student_data' not in st.session_state:
    st.session_state.student_data = None
if 'separation_data' not in st.session_state:
    st.session_state.separation_data = None
if 'assignment_file' not in st.session_state:
    st.session_state.assignment_file = None

# Step 1: 학생자료 업로드
if st.session_state.step == 1:
    st.header("1단계: 학생자료 업로드")
    st.info("학생자료.xlsx 파일을 업로드해주세요.")
    
    uploaded_file = st.file_uploader(
        "학생자료.xlsx 파일 선택",
        type=['xlsx'],
        key="student_file"
    )
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state.student_data = df
            st.success(f"✅ 학생자료 업로드 완료! (총 {len(df)}명)")
            st.dataframe(df.head(), use_container_width=True)
            
            if st.button("다음 단계로", type="primary"):
                st.session_state.step = 2
                st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")

# Step 2: 분리명부 업로드
elif st.session_state.step == 2:
    st.header("2단계: 분리명부 업로드")
    st.info("separation.xlsx 파일을 업로드해주세요.")
    
    # 학생자료 미리보기
    if st.session_state.student_data is not None:
        with st.expander("업로드된 학생자료 확인"):
            st.dataframe(st.session_state.student_data, use_container_width=True)
    
    uploaded_file = st.file_uploader(
        "separation.xlsx 파일 선택",
        type=['xlsx'],
        key="separation_file"
    )
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state.separation_data = df
            st.success(f"✅ 분리명부 업로드 완료! (총 {len(df)}개 규칙)")
            st.dataframe(df.head(), use_container_width=True)
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("이전 단계로"):
                    st.session_state.step = 1
                    st.rerun()
            with col2:
                if st.button("배정표 생성", type="primary"):
                    with st.spinner("배정표 생성 중..."):
                        try:
                            # 임시 파일로 저장
                            student_temp = "temp_student_data.xlsx"
                            separation_temp = "temp_separation_data.xlsx"
                            
                            st.session_state.student_data.to_excel(student_temp, index=False)
                            st.session_state.separation_data.to_excel(separation_temp, index=False)
                            
                            # 기존 파일명 백업
                            original_student = "학생자료.xlsx"
                            original_separation = "separation.xlsx"
                            
                            # 기존 파일이 있으면 백업
                            if os.path.exists(original_student):
                                shutil.copy(original_student, original_student + ".bak")
                            if os.path.exists(original_separation):
                                shutil.copy(original_separation, original_separation + ".bak")
                            
                            # 임시 파일을 원래 이름으로 복사
                            shutil.copy(student_temp, original_student)
                            shutil.copy(separation_temp, original_separation)
                            
                            # create_final_assignment.py 실행
                            # subprocess로 실행하여 독립적으로 처리
                            import subprocess
                            result = subprocess.run(
                                [sys.executable, "create_final_assignment.py"],
                                capture_output=True,
                                text=True,
                                encoding='utf-8',
                                errors='ignore'
                            )
                            
                            if result.returncode != 0:
                                st.error(f"배정표 생성 중 오류 발생:\n{result.stderr}")
                                raise Exception(f"배정표 생성 실패: {result.stderr}")
                            
                            output_file = "반편성_배정표.xlsx"
                            
                            # 임시 파일 삭제
                            if os.path.exists(student_temp):
                                os.remove(student_temp)
                            if os.path.exists(separation_temp):
                                os.remove(separation_temp)
                            
                            st.session_state.assignment_file = output_file
                            st.session_state.step = 3
                            st.success("✅ 배정표 생성 완료!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"배정표 생성 오류: {str(e)}")
                            st.exception(e)
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")

# Step 3: 배정표 다운로드 및 완료 파일 업로드
elif st.session_state.step == 3:
    st.header("3단계: 배정표 다운로드 및 완료 파일 업로드")
    st.info("생성된 배정표를 다운로드하여 수동 배정을 완료한 후, 완료 파일을 업로드해주세요.")
    
    # 배정표 다운로드
    if st.session_state.assignment_file and os.path.exists(st.session_state.assignment_file):
        with open(st.session_state.assignment_file, "rb") as f:
            st.download_button(
                label="📥 반편성_배정표.xlsx 다운로드",
                data=f.read(),
                file_name="반편성_배정표.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    uploaded_file = st.file_uploader(
        "반편성_완료.xlsx 파일 선택 (수동 배정 완료 후)",
        type=['xlsx'],
        key="completed_file_uploader"
    )
    
    if uploaded_file is not None:
        try:
            # 임시 파일로 저장
            completed_temp = "temp_completed_data.xlsx"
            with open(completed_temp, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            st.success("✅ 완료 파일 업로드 완료!")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("이전 단계로"):
                    st.session_state.step = 2
                    st.rerun()
            with col2:
                if st.button("최종 출력서식 생성", type="primary"):
                    st.session_state.completed_file = completed_temp
                    st.session_state.step = 4
                    st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")

# Step 4: 최종 출력서식 생성
elif st.session_state.step == 4:
    st.header("4단계: 최종 출력서식 생성")
    
    # 출력서식 템플릿 확인
    template_file = "출력서식.xlsx"
    if not os.path.exists(template_file):
        st.error(f"❌ 출력서식.xlsx 템플릿 파일이 없습니다. 현재 디렉토리에 '{template_file}' 파일을 추가해주세요.")
    else:
        completed_temp = st.session_state.get('completed_file', 'temp_completed_data.xlsx')
        
        if st.button("최종 출력서식 생성", type="primary"):
            with st.spinner("최종 출력서식 생성 중..."):
                try:
                    # 임시 파일 경로
                    student_temp = "temp_student_data.xlsx"
                    
                    # 학생자료 임시 저장
                    if st.session_state.student_data is not None:
                        st.session_state.student_data.to_excel(student_temp, index=False)
                    
                    # 기존 파일명 백업
                    original_completed = "반편성_완료.xlsx"
                    original_student = "학생자료.xlsx"
                    
                    # 기존 파일이 있으면 백업
                    if os.path.exists(original_completed):
                        shutil.copy(original_completed, original_completed + ".bak")
                    if os.path.exists(original_student):
                        shutil.copy(original_student, original_student + ".bak")
                    
                    # 임시 파일을 원래 이름으로 복사
                    shutil.copy(completed_temp, original_completed)
                    shutil.copy(student_temp, original_student)
                    
                    # fill_output_format.py 실행
                    # subprocess로 실행하여 독립적으로 처리
                    result = subprocess.run(
                        [sys.executable, "fill_output_format.py"],
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        errors='ignore'
                    )
                    
                    if result.returncode != 0:
                        st.error(f"출력서식 생성 중 오류 발생:\n{result.stderr}")
                        raise Exception(f"출력서식 생성 실패: {result.stderr}")
                    
                    output_file = "출력서식_완료.xlsx"
                    
                    # 임시 파일 삭제
                    if os.path.exists(student_temp):
                        os.remove(student_temp)
                    
                    st.success("✅ 최종 출력서식 생성 완료!")
                    
                    # 다운로드 버튼
                    if os.path.exists(output_file):
                        with open(output_file, "rb") as f:
                            st.download_button(
                                label="📥 출력서식_완료.xlsx 다운로드",
                                data=f.read(),
                                file_name="출력서식_완료.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    
                    if st.button("처음부터 다시 시작"):
                        # 세션 상태 초기화
                        for key in list(st.session_state.keys()):
                            del st.session_state[key]
                        st.session_state.step = 1
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"출력서식 생성 오류: {str(e)}")
                    st.exception(e)
        else:
            st.info("위의 '최종 출력서식 생성' 버튼을 클릭하세요.")

# 사이드바
with st.sidebar:
    st.header("진행 상황")
    steps = [
        "1. 학생자료 업로드",
        "2. 분리명부 업로드",
        "3. 배정표 다운로드",
        "4. 최종 출력서식"
    ]
    
    for i, step_name in enumerate(steps, 1):
        if i < st.session_state.step:
            st.success(f"✅ {step_name}")
        elif i == st.session_state.step:
            st.info(f"🔄 {step_name} (진행 중)")
        else:
            st.write(f"⏳ {step_name}")
    
    st.markdown("---")
    st.markdown("### 사용 방법")
    st.markdown("""
    1. **학생자료.xlsx** 업로드
    2. **separation.xlsx** 업로드
    3. 생성된 **반편성_배정표.xlsx** 다운로드
    4. 수동 배정 완료 후 **반편성_완료.xlsx** 업로드
    5. **출력서식_완료.xlsx** 다운로드
    """)
