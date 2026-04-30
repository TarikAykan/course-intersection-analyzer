from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass
class AnalysisResult:
    result_df: pd.DataFrame
    total_count: int


def extract_unique_courses(df: pd.DataFrame, course_col: str) -> list[str]:
    courses = (
        df[course_col]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s != ""]
        .drop_duplicates()
        .sort_values()
    )
    return courses.tolist()


def _build_student_key(row: pd.Series, student_no_col: str, full_name_col: str) -> str:
    student_no = str(row.get(student_no_col, "")).strip() if student_no_col else ""
    full_name = str(row.get(full_name_col, "")).strip().lower()

    if student_no and student_no.lower() != "nan":
        return f"NO::{student_no}"
    if full_name and full_name.lower() != "nan":
        return f"AD::{full_name}"
    return ""


def find_students_with_two_courses(
    df: pd.DataFrame,
    student_no_col: str,
    first_name_col: str,
    last_name_col: str,
    course_col: str,
    first_course: str,
    second_course: str,
) -> AnalysisResult:
    working_df = df.copy()
    if student_no_col:
        working_df[student_no_col] = working_df[student_no_col].fillna("").astype(str).str.strip()
    working_df[first_name_col] = working_df[first_name_col].fillna("").astype(str).str.strip()
    working_df[last_name_col] = working_df[last_name_col].fillna("").astype(str).str.strip()
    working_df[course_col] = working_df[course_col].fillna("").astype(str).str.strip()
    working_df["_full_name"] = (
        working_df[first_name_col].str.cat(working_df[last_name_col], sep=" ").str.strip()
    )

    working_df = working_df[working_df[course_col] != ""]
    if working_df.empty:
        return AnalysisResult(result_df=pd.DataFrame(), total_count=0)

    working_df["_student_key"] = working_df.apply(
        _build_student_key,
        axis=1,
        student_no_col=student_no_col,
        full_name_col="_full_name",
    )
    working_df = working_df[working_df["_student_key"] != ""]

    selected_courses = {first_course, second_course}
    grouped = working_df.groupby("_student_key", dropna=False)

    rows: list[dict[str, str]] = []
    for _, group in grouped:
        student_courses = set(group[course_col].tolist())
        if not selected_courses.issubset(student_courses):
            continue

        first_row = group.iloc[0]
        student_no_value = str(first_row.get(student_no_col, "")).strip() if student_no_col else ""
        full_name_value = str(first_row["_full_name"]).strip()
        relevant_courses = [course for course in [first_course, second_course] if course in student_courses]

        rows.append(
            {
                "Ogrenci No": "" if student_no_value.lower() == "nan" else student_no_value,
                "Ad Soyad": "" if full_name_value.lower() == "nan" else full_name_value,
                "Aldigi Ilgili Dersler": ", ".join(relevant_courses),
            }
        )

    result_df = pd.DataFrame(rows)
    if not result_df.empty:
        result_df = result_df.sort_values(by=["Ogrenci No", "Ad Soyad"], kind="stable").reset_index(drop=True)

    return AnalysisResult(result_df=result_df, total_count=len(result_df))
