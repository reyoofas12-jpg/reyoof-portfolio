import pandas as pd

# ===== تحميل البيانات الكاملة =====
df_full = pd.read_excel("IHR_مدراء_المشاريع_v2.xlsx",
                        sheet_name="التفاصيل الكاملة",
                        header=1)
df_full = df_full.dropna(how="all")

# ===== نفس كودك الأصلي + شيت التفاصيل =====
df = df_full.copy()
df["Contract End Date"] = pd.to_datetime(df["Contract End Date"], errors="coerce")
today = pd.Timestamp.today()

with pd.ExcelWriter("IHR_Results.xlsx", engine="openpyxl") as writer:

    # ✅ الجديد — التفاصيل الكاملة أول شيت
    df_full.to_excel(writer, sheet_name="التفاصيل الكاملة", index=False)

    # باقي الشيتات — نفس كودك بالضبط
    df["مدير المشروع"].value_counts().reset_index().rename(
        columns={"مدير المشروع": "مدير المشروع", "count": "عدد"}).to_excel(writer, sheet_name="المدراء", index=False)

    df["Nationality"].value_counts().reset_index().rename(
        columns={"Nationality": "الجنسية", "count": "عدد"}).to_excel(writer, sheet_name="الجنسيات", index=False)

    df["Contract Extension (Y/N)"].value_counts().reset_index().rename(
        columns={"Contract Extension (Y/N)": "التمديد", "count": "عدد"}).to_excel(writer, sheet_name="تمديد العقود", index=False)

    df["Project"].value_counts().reset_index().rename(
        columns={"Project": "المشروع", "count": "عدد"}).to_excel(writer, sheet_name="المشاريع", index=False)

    df["Role"].value_counts().reset_index().rename(
        columns={"Role": "الدور", "count": "عدد"}).to_excel(writer, sheet_name="الأدوار", index=False)

    df["Contract Duration"].value_counts().reset_index().rename(
        columns={"Contract Duration": "مدة العقد", "count": "عدد"}).to_excel(writer, sheet_name="مدة العقود", index=False)

    soon = df[df["Contract End Date"].between(today, today + pd.DateOffset(months=3))]
    soon[["مدير المشروع", "ID", "Role", "Contract End Date"]].to_excel(writer, sheet_name="عقود تنتهي قريباً", index=False)

    expired = df[df["Contract End Date"] < today]
    expired[["مدير المشروع", "ID", "Role", "Contract End Date"]].to_excel(writer, sheet_name="عقود منتهية", index=False)

print("✅ تم!")