# -*- coding: utf-8 -*-
"""
반편성 배정 프로그램 — 4반/3반 편성 선택

처음 실행 시 4반 편성(자유 배정) 또는 3반 편성(선생님 연속 지도 배제)을 선택한 뒤
기존과 동일한 흐름으로 진행합니다.
"""
import streamlit as st
import pandas as pd
import sys
import os
import shutil
import subprocess

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
    os.chdir(application_path)
else:
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
    st.session_state.step = 0
if 'assign_mode' not in st.session_state:
    st.session_state.assign_mode = "4반"
if 'student_data' not in st.session_state:
    st.session_state.student_data = None
if 'separation_data' not in st.session_state:
    st.session_state.separation_data = None
if 'assignment_file' not in st.session_state:
    st.session_state.assignment_file = None

# Step 0: 4반/3반 편성 선택
if st.session_state.step == 0:
    st.header("0단계: 편성 방식 선택")
    st.info("4반 편성(자유 배정) 또는 3반 편성(선생님 연속 지도 배제) 중 하나를 선택하세요.")

    mode = st.radio(
        "편성 방식",
        ["4반", "3반"],
        format_func=lambda x: (
            "**4반 편성** — 1,2,3,4반 → A,B,C,D 자유 배정 (기존과 동일)"
            if x == "4반" else
            "**3반 편성** — 1→B,C,D / 2→A,C,D / 3→A,B,D / 4→A,B,C (담당 선생님 연속 지도 배제)"
        ),
        horizontal=True,
        key="mode_radio"
    )
    st.session_state.assign_mode = mode

    if st.button("다음 단계로", type="primary"):
        st.session_state.step = 1
        st.rerun()

# Step 1: 학생자료 업로드
elif st.session_state.step == 1:
    st.header("1단계: 학생자료 업로드")
    st.info("학생자료.xlsx 파일을 업로드해주세요. 양식을 다운받아 수정 후 첨부할 수 있습니다.")
    st.caption(f"현재 선택: **{st.session_state.assign_mode} 편성**")

    student_template = "학생자료_서식.xlsx"
    if os.path.exists(student_template):
        with open(student_template, "rb") as f:
            st.download_button(
                label="📥 학생자료 양식 다운로드 (학생자료_서식.xlsx)",
                data=f.read(),
                file_name="학생자료_서식.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_student_template"
            )
    else:
        st.caption(f"양식 파일 `{student_template}` 이 없습니다.")

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

            col1, col2 = st.columns(2)
            with col1:
                if st.button("이전 단계로"):
                    st.session_state.step = 0
                    st.rerun()
            with col2:
                if st.button("다음 단계로", type="primary"):
                    st.session_state.step = 2
                    st.rerun()
        except Exception as e:
            st.error(f"파일 읽기 오류: {str(e)}")
    else:
        if st.button("이전 단계로"):
            st.session_state.step = 0
            st.rerun()

# Step 2: 분리명부 업로드
elif st.session_state.step == 2:
    st.header("2단계: 분리명부 업로드")
    st.info("separation.xlsx 파일을 업로드해주세요. 양식을 다운받아 수정 후 첨부할 수 있습니다.")
    st.caption(f"현재 선택: **{st.session_state.assign_mode} 편성**")

    separation_template = "separation_서식.xlsx"
    if os.path.exists(separation_template):
        with open(separation_template, "rb") as f:
            st.download_button(
                label="📥 분리명부 양식 다운로드 (separation_서식.xlsx)",
                data=f.read(),
                file_name="separation_서식.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_separation_template"
            )
    else:
        st.caption(f"양식 파일 `{separation_template}` 이 없습니다.")

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
                            student_temp = "temp_student_data.xlsx"
                            separation_temp = "temp_separation_data.xlsx"

                            st.session_state.student_data.to_excel(student_temp, index=False)
                            st.session_state.separation_data.to_excel(separation_temp, index=False)

                            original_student = "학생자료.xlsx"
                            original_separation = "separation.xlsx"

                            if os.path.exists(original_student):
                                shutil.copy(original_student, original_student + ".bak")
                            if os.path.exists(original_separation):
                                shutil.copy(original_separation, original_separation + ".bak")

                            shutil.copy(student_temp, original_student)
                            shutil.copy(separation_temp, original_separation)

                            # 환경변수로 ASSIGN_MODE 전달
                            env = os.environ.copy()
                            env["ASSIGN_MODE"] = st.session_state.assign_mode

                            result = subprocess.run(
                                [sys.executable, "create_final_assignment.py"],
                                capture_output=True,
                                text=True,
                                encoding='utf-8',
                                errors='ignore',
                                env=env
                            )

                            if result.returncode != 0:
                                error_msg = result.stderr if result.stderr else result.stdout
                                st.error(f"배정표 생성 중 오류 발생:\n{error_msg}")
                                # stdout도 표시 (미배정 학생 정보 등)
                                if result.stdout and "배정되지 않은 학생" in result.stdout:
                                    st.warning("⚠️ 자세한 오류 정보:")
                                    st.code(result.stdout, language="text")
                                raise Exception(f"배정표 생성 실패: {error_msg}")

                            output_file = "반편성_배정표.xlsx"

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
    else:
        if st.button("이전 단계로"):
            st.session_state.step = 1
            st.rerun()

# Step 3: 배정표 다운로드 및 완료 파일 업로드
elif st.session_state.step == 3:
    st.header("3단계: 배정표 다운로드 및 완료 파일 업로드")
    st.info("생성된 배정표를 다운로드하여 수동 배정을 완료한 후, 완료 파일을 업로드해주세요.")
    st.caption(f"현재 선택: **{st.session_state.assign_mode} 편성**")

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
    else:
        if st.button("이전 단계로"):
            st.session_state.step = 2
            st.rerun()

# Step 4: 최종 출력서식 생성
elif st.session_state.step == 4:
    st.header("4단계: 최종 출력서식 생성")
    st.caption(f"현재 선택: **{st.session_state.assign_mode} 편성**")

    template_file = "출력서식.xlsx"
    if not os.path.exists(template_file):
        st.error(f"❌ 출력서식.xlsx 템플릿 파일이 없습니다. 현재 디렉토리에 '{template_file}' 파일을 추가해주세요.")
    else:
        completed_temp = st.session_state.get('completed_file', 'temp_completed_data.xlsx')

        if st.button("최종 출력서식 생성", type="primary"):
            with st.spinner("최종 출력서식 생성 중..."):
                try:
                    student_temp = "temp_student_data.xlsx"

                    if st.session_state.student_data is not None:
                        st.session_state.student_data.to_excel(student_temp, index=False)

                    original_completed = "반편성_완료.xlsx"
                    original_student = "학생자료.xlsx"

                    if os.path.exists(original_completed):
                        shutil.copy(original_completed, original_completed + ".bak")
                    if os.path.exists(original_student):
                        shutil.copy(original_student, original_student + ".bak")

                    shutil.copy(completed_temp, original_completed)
                    shutil.copy(student_temp, original_student)

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

                    if os.path.exists(student_temp):
                        os.remove(student_temp)

                    st.success("✅ 최종 출력서식 생성 완료!")

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
                        for key in list(st.session_state.keys()):
                            del st.session_state[key]
                        st.session_state.step = 0
                        st.session_state.assign_mode = "4반"
                        st.rerun()

                except Exception as e:
                    st.error(f"출력서식 생성 오류: {str(e)}")
                    st.exception(e)
        else:
            st.info("위의 '최종 출력서식 생성' 버튼을 클릭하세요.")
            if st.button("이전 단계로"):
                st.session_state.step = 3
                st.rerun()

# 사이드바
with st.sidebar:
    st.header("진행 상황")
    steps = [
        "0. 4반/3반 선택",
        "1. 학생자료 업로드",
        "2. 분리명부 업로드",
        "3. 배정표 다운로드",
        "4. 최종 출력서식"
    ]

    for i, step_name in enumerate(steps):
        if i < st.session_state.step:
            st.success(f"✅ {step_name}")
        elif i == st.session_state.step:
            st.info(f"🔄 {step_name} (진행 중)")
        else:
            st.write(f"⏳ {step_name}")

    st.markdown("---")
    if st.session_state.step >= 0:
        st.caption(f"**편성:** {st.session_state.assign_mode}")
    st.markdown("---")
    st.markdown("### 사용 방법")
    st.markdown("""
    0. **4반/3반 편성** 선택
    1. **학생자료.xlsx** 업로드
    2. **separation.xlsx** 업로드
    3. **반편성_배정표.xlsx** 다운로드 후 수동 배정
    4. **반편성_완료.xlsx** 업로드
    5. **출력서식_완료.xlsx** 다운로드
    """)
