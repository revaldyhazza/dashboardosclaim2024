import streamlit as st
import plotly.express as px
import pandas as pd
import matplotlib.pyplot as plt
import plotly.figure_factory as ff
import os
import io
import warnings
warnings.filterwarnings('ignore')
st.set_page_config(page_title="DB OS Claim 2024", page_icon=":bar_chart:",layout="wide")

# Tampilkan judul langsung di bawah gambar
st.markdown("# 📊 Dashboard OS Claim 2024")
st.markdown("##### 📍 For better experience, please click left side arrow on the screen < after using the filter")

# Load default data
os.chdir(r"/Users/revaldyhazzadaniswara/Downloads")
df = pd.read_csv("OS Gabung Fix.csv", encoding="ISO-8859-1").drop(columns=['Unnamed: 0', 'Unnamed: 11'], errors='ignore').dropna(axis=1, how='all').drop(columns=['Unnamed: 11'], errors='ignore').dropna(axis=1, how='all')

# Show preview
st.dataframe(df.head())

# File uploader
fl = st.file_uploader(":file_folder: If there is any data that you want to analyze with the same format as our data above, please upload here", type=["csv", "txt", "xlsx", "xls"])

# Load uploaded file if available
if fl is not None:
    filename = fl.name
    st.write(f"Uploaded file: {filename}")
    df = pd.read_csv(fl, encoding="ISO-8859-1")
    
    # Show preview of uploaded file
    st.write("### Uploaded Data Preview")
    st.dataframe(df.head())

st.logo("/Users/revaldyhazzadaniswara/Documents/Screenshots/Logo Askrindo.png")
st.sidebar.header("Choose your filter:")

# Filter CoB
cob = st.sidebar.multiselect("Pick Classification", df["CoB"].unique())
df2 = df[df["CoB"].isin(cob)] if cob else df.copy()

# Filter Type of Policy
jenis = st.sidebar.multiselect("Pick Type of Policy", df2["Type"].unique())
df3 = df2[df2["Type"].isin(jenis)] if jenis else df2.copy()

# Filter Business
bisnis = st.sidebar.multiselect("Pick Business", df3["Business"].unique())
df4 = df3[df3["Business"].isin(bisnis)] if bisnis else df3.copy()

# Mapping angka bulan ke nama bulan
month_mapping = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}
df4["Month"] = df4["Month"].replace(month_mapping)

# Urutkan kategori bulan agar tidak acak
order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
df4["Month"] = pd.Categorical(df4["Month"], categories=order, ordered=True)

# Sidebar multiselect dengan nama bulan (tanpa default semua kepilih)
bulan = st.sidebar.multiselect("Pick Month", df4["Month"].unique())

# Filter berdasarkan bulan
df5 = df4[df4["Month"].isin(bulan)] if bulan else df4.copy()
df5["Gross Claim (In Million)"] = pd.to_numeric(df5["Gross Claim (In Million)"], errors="coerce")

filter_text = "**Filters Applied:**\n"
if cob:
    filter_text += f"- **CoB:** {', '.join(cob)}\n"
if jenis:
    filter_text += f"- **Type of Policy:** {', '.join(jenis)}\n"
if bisnis:
    filter_text += f"- **Business:** {', '.join(bisnis)}\n"
if bulan:
    filter_text += f"- **Month:** {', '.join(bulan)}\n"

if filter_text == "**Filters Applied:**\n":
    filter_text = "No filters applied."

st.markdown(filter_text)

# Hitung kembali setelah menangani NaN
total_severity = df5["Gross Claim (In Million)"].sum()
total_frequency = df5["Gross Claim (In Million)"].count()
rata_rata_klaim = total_severity / total_frequency if total_frequency > 0 else 0

# Format angka dengan titik sebagai pemisah ribuan
formatted_total_severity = f"{total_severity:,.0f}".replace(",", ".")
formatted_total_frequency = f"{total_frequency:,.0f}".replace(",", ".")
formatted_rata_rata_klaim = f"{rata_rata_klaim:,.2f}".replace(",", ".")

# Layout
st.markdown("""
    <style>
        .metric-box {
            background-color: #D67230;
            padding: 15px;
            text-align: center;
            border-radius: 5px;
            font-weight: bold;
            color: white;
            font-size: 24px;
        }
        .value {
            font-size: 24px;
            color: black;
            font-weight: bold;
            text-align: center;
            background-color: white;
            padding: 10px;
            border-radius: 5px;
            margin-top: 5px;
        }
    </style>
""", unsafe_allow_html=True)


col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="metric-box">Total Severity Claim</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="value">{formatted_total_severity}</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="metric-box">Total Frequency Claim</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="value">{formatted_total_frequency}</div>', unsafe_allow_html=True)

with col3:
    st.markdown('<div class="metric-box">Average Claim</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="value">{formatted_rata_rata_klaim}</div>', unsafe_allow_html=True)

# Filter berdasarkan CoB
df5 = df5.copy()
if cob:
    df5 = df5[df5["CoB"].isin(cob)]
if jenis:
    df5 = df5[df5["Type"].isin(jenis)]
if bisnis:
    df5 = df5[df5["Business"].isin(bisnis)]
if bulan:
    df5 = df5[df5["Month"].isin(bulan)]

# **Tampilkan Pivot Table**
order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
df5["Month"] = pd.Categorical(df5["Month"], categories=order, ordered=True)

st.markdown("### 📌 Monthly Gross Claim by Type Summary")

# Buat Pivot Table
pivot_df = pd.pivot_table(df5, values="Gross Claim (In Million)", 
                          index=["CoB", "Type"], columns="Month", aggfunc="sum")

# Hapus baris dengan semua nilai nol
pivot_df = pivot_df.loc[pivot_df.sum(axis=1) != 0]

# Tampilkan di Streamlit dengan lebar otomatis dan formatting
with st.expander("Click Here"):
    st.write(pivot_df.style.format("{:,.0f}").background_gradient(cmap="Blues"))


st.markdown("### 📌 Monthly Gross Claim by Business Summary")

# Buat Pivot Table
pivot_df = pd.pivot_table(df5, values="Gross Claim (In Million)", 
                          index=["CoB", "Business"], columns="Month", aggfunc="sum")

# Hapus baris dengan semua nilai nol
pivot_df = pivot_df.loc[pivot_df.sum(axis=1) != 0]

# Tampilkan di Streamlit dengan lebar otomatis dan formatting
with st.expander("Click Here"):
    st.write(pivot_df.style.format("{:,.0f}").background_gradient(cmap="Blues"))


# **Pivot Table untuk Total Gross Claim per Bulan**
st.markdown("### :calendar: Monthly Gross Claim Summary")
pivot_df = pd.pivot_table(df5, values="Gross Claim (In Million)", index=["CoB"], columns="Month", aggfunc="sum")

# **Tampilkan pivot table dengan warna gradasi biru**
with st.expander("Click Here"):
    st.write(pivot_df.style.format("{:,.0f}").background_gradient(cmap="Blues"))

# Ubah kolom "Gross Claim (In Million)" menjadi numerik, paksa error sebagai NaN jika ada data non-angka
df5["Gross Claim (In Million)"] = pd.to_numeric(df5["Gross Claim (In Million)"], errors="coerce")

# Kelompokkan berdasarkan CoB dan jumlahkan klaim
cob_severity = df5.groupby(by=["CoB"], as_index=False)["Gross Claim (In Million)"].sum()

cob_frequency = df5.groupby(by=["CoB"], as_index=False)["Gross Claim (In Million)"].count()

import plotly.express as px

# Format angka lebih sederhana (T untuk triliun, M untuk juta)
def simplify_number(value):
    if value >= 1e12:  # Triliun
        return f"{value / 1e12:.1f} T"
    elif value >= 1e9:  # Miliar
        return f"{value / 1e9:.1f} B"
    elif value >= 1e6:  # Juta
        return f"{value / 1e6:.1f} M"
    else:  # Ratusan ribu atau lebih kecil
        return f"{value:,.0f}"

# Apply ke semua nilai dalam kolom "Gross Claim (In Million)"
cob_severity["Formatted Claim"] = cob_severity["Gross Claim (In Million)"].apply(simplify_number)

with col1:
    st.markdown(
    """
    <style>
        .custom-title {
            text-align: center;
            margin-bottom: 40px; /* Sesuaikan angka ini */
            margin-top: 20px; /* Atur naik-turun */
        }
    </style>
    <h4 class="custom-title">Severity Based on CoB</h4>
    """,
    unsafe_allow_html=True
)

    st.subheader("")
    fig = px.bar(
        cob_severity, 
        x="CoB", 
        y="Gross Claim (In Million)", 
        text=cob_severity["Formatted Claim"],  # Gunakan teks yang sudah disederhanakan
        template="seaborn",
        color_discrete_sequence=["#FFF5EB", "#FED98E"]
    )
    fig.update_layout(
    width=200,  # Atur lebar plot
    height=340, # Atur tinggi plot
    margin=dict(l=10, r=10, t=5, b=5)  # Perkecil margin
)

    fig.update_traces(textposition="outside")  # Letakkan teks di luar batang
    st.plotly_chart(fig, use_container_width=True)  # Perbesar tinggi agar lebih proporsional

with col2:
    st.markdown(
    """
    <style>
        .custom-title {
            text-align: center;
            margin-bottom: -20px; /* Sesuaikan angka ini */
            position: relative;
            top: 20px; /* Atur naik-turun */
        }
    </style>
    <h4 class="custom-title">Severity Based on CoB</h4>
    """,
    unsafe_allow_html=True
)


    fig = px.pie(
        cob_severity, 
        values="Gross Claim (In Million)", 
        names="CoB", 
        hole=0.5,  # Membuat donut chart
        color_discrete_sequence=["#FFF5EB", "#FED98E"]  # Warna oranye bertingkat
    )

    fig.update_layout(
    width=400,  # Lebar keseluruhan
        height=250, # Tinggi keseluruhan
        margin=dict(l=10, r=10, t=50, b=10),
        font=dict(size=12),
        legend=dict(
            font=dict(size=11),
            x=1,  # Geser legend ke kanan
            y=0.5,
            xanchor="left",
            yanchor="middle",
        ),  # Perkecil margin
)

    # Menampilkan label nama CoB di luar chart
    
    st.plotly_chart(fig, width = 5)

with col3:
    st.markdown(
    """
    <style>
        .custom-title {
            text-align: center;
            margin-bottom: -20px; /* Sesuaikan angka ini */
            position: relative;
            top: 20px; /* Atur naik-turun */
        }
    </style>
    <h4 class="custom-title">Frequency Based on CoB</h4>
    """,
    unsafe_allow_html=True
)


    fig = px.pie(
        cob_frequency, 
        values="Gross Claim (In Million)", 
        names="CoB", 
        hole=0.5,  # Membuat donut chart
        color_discrete_sequence=["#B3D1FF", "#00205F"]
    )

    fig.update_layout(
    width=400,  # Lebar keseluruhan
        height=250, # Tinggi keseluruhan
        margin=dict(l=10, r=10, t=50, b=10),
        font=dict(size=12),
        legend=dict(
            font=dict(size=12),
            x=1,  # Geser legend ke kanan
            y=0.5,
            xanchor="left",
            yanchor="middle",
        ),  # Perkecil margin
    )

    st.plotly_chart(fig, width = 5)

# Agregasi data untuk pie chart (ambil 5 teratas)
CoC_counts = df5["Cause of Claim"].value_counts().nlargest(5).reset_index()
CoC_counts.columns = ["Cause of Claim", "Count"]

CoC_sev = df5.groupby("Cause of Claim")["Gross Claim (In Million)"].sum().nlargest(5).reset_index()
CoC_sev.columns=["Cause of Claim", "Severity"]

# Buat dua kolom dengan lebar yang sama
col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        <style>
            .custom-title {
                text-align: center;
                margin-bottom: -20px;
                position: relative;
                top: 10px;
            }
        </style>
        <h5 class="custom-title">Top 5 Frequency by Claim Cause</h5>
        """,
        unsafe_allow_html=True
    )

    # Plotly Pie Chart untuk Frequency
    colors = ["#FFF5EB", "#FED98E", "#FE9929", "#D95F0E", "#993404"]  # Oranye bertingkat
    fig1 = px.pie(CoC_counts, names="Cause of Claim", values="Count", hole=0.3, color_discrete_sequence=colors)

    # Atur ukuran chart & font agar seimbang
    fig1.update_layout(
        width=400,  # Lebar keseluruhan
        height=250, # Tinggi keseluruhan
        margin=dict(l=10, r=10, t=50, b=10),
        font=dict(size=12),
        legend=dict(
            title="Claim Causes",
            font=dict(size=12),
            x=1,  # Geser legend ke kanan
            y=0.5,
            xanchor="left",
            yanchor="middle",
        ),
    )

    # Tampilkan di Streamlit
    st.plotly_chart(fig1)

with col2:
    st.markdown(
        """
        <style>
            .custom-title {
                text-align: center;
                margin-bottom: 0px;
                position: relative;
                top: 10px;
            </style>
        <h5 class="custom-title">Top 5 Severity by Claim Cause</h5>
        """,
        unsafe_allow_html=True
    )

    # Plotly Pie Chart untuk Severity
    fig2 = px.pie(CoC_sev, names="Cause of Claim", values="Severity", hole=0.3, color_discrete_sequence=colors)

    # Atur ukuran chart & font agar seimbang
    fig2.update_layout(
        width=400,  # Lebar keseluruhan
        height=250, # Tinggi keseluruhan
        margin=dict(l=10, r=10, t=50, b=10),
        font=dict(size=12),
        legend=dict(
            title="Claim Causes",
            font=dict(size=12),
            x=1,  # Geser legend ke kanan
            y=0.5,
            xanchor="left",
            yanchor="middle",
        ),
    )

    # Tampilkan di Streamlit
    st.plotly_chart(fig2)

# Agregasi data untuk pie chart (ambil 5 teratas)
KC_counts = df5["KC"].value_counts().nlargest(5).reset_index()
KC_counts.columns = ["KC", "Count"]

KC_sev = df5.groupby("KC")["Gross Claim (In Million)"].sum().nlargest(5).reset_index()
KC_sev.columns=["KC", "Severity"]

# Buat dua kolom dengan lebar yang sama
col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        <style>
            .custom-title {
                text-align: center;
                margin-bottom: -20px;
                position: relative;
                top: 10px;
            }
        </style>
        <h4 class="custom-title">Top 5 Frequency by KC</h4>
        """,
        unsafe_allow_html=True
    )

    # Plotly Pie Chart untuk Frequency
    colors = colors = ["#00205F", "#004080", "#0074CC", "#66A3FF", "#B3D1FF"] 
    fig1 = px.pie(KC_counts, names="KC", values="Count", hole=0.3, color_discrete_sequence=colors)

    # Atur ukuran chart & font agar seimbang
    fig1.update_layout(
        width=400,  # Lebar keseluruhan
        height=250, # Tinggi keseluruhan
        margin=dict(l=10, r=10, t=50, b=10),
        font=dict(size=12),
        legend=dict(
            title="KC",
            font=dict(size=12),
            x=1,  # Geser legend ke kanan
            y=0.5,
            xanchor="left",
            yanchor="middle",
        ),
    )

    # Tampilkan di Streamlit
    st.plotly_chart(fig1)

with col2:
    st.markdown(
        """
        <style>
            .custom-title {
                text-align: center;
                margin-bottom: 0px;
                position: relative;
                top: 10px;
            </style>
        <h4 class="custom-title">Top 5 Severity by KC</h4>
        """,
        unsafe_allow_html=True
    )

    # Plotly Pie Chart untuk Severity
    colors = ["#00205F", "#004080", "#0074CC", "#66A3FF", "#B3D1FF"]
    fig2 = px.pie(KC_sev, names="KC", values="Severity", hole=0.3, color_discrete_sequence=colors)

    # Atur ukuran chart & font agar seimbang
    fig2.update_layout(
        width=400,  # Lebar keseluruhan
        height=250, # Tinggi keseluruhan
        margin=dict(l=10, r=10, t=50, b=10),
        font=dict(size=12),
        legend=dict(
            title="KC",
            font=dict(size=12),
            x=1,  # Geser legend ke kanan
            y=0.5,
            xanchor="left",
            yanchor="middle",
        ),
    )

    # Tampilkan di Streamlit
    st.plotly_chart(fig2)

# Agregasi data untuk pie chart (ambil 5 teratas)
Claimant_counts = df5["Claimant"].value_counts().nlargest(10).reset_index()
Claimant_counts.columns = ["Claimant", "Count"]

Claimant_sev = df5.groupby("Claimant")["Gross Claim (In Million)"].sum().nlargest(10).reset_index()
Claimant_sev.columns=["Claimant", "Severity"]

Claimant_sev["Severity"] = Claimant_sev["Severity"].apply(simplify_number)

col1, col2 = st.columns(2)
with col1:
    st.markdown(
        """
        <style>
            .custom-title {
                text-align: center;
                margin-bottom: -20px;
                position: relative;
                top: 10px;
            }
        </style>
        <h4 class="custom-title">Top 10 Frequency by Claimant</h4>
        """,
        unsafe_allow_html=True
    )

    # Horizontal Bar Plot untuk Frequency
    colors = [
    "#004d00",  
    "#006600",  
    "#008000",  
    "#009900",  
    "#00b300",  
    "#00cc00",  
    "#33ff33",  
    "#80ff80",  
    "#ccffcc",  
    "#ffffff"   
]
    fig1 = px.bar(Claimant_counts, 
                  x="Count", 
                  y="Claimant", 
                  orientation="h", 
                  text="Count", 
                  color="Claimant", 
                  color_discrete_sequence=colors)

    # Atur layout untuk memperbaiki tampilan
    fig1.update_layout(
    width=900,  # Lebar lebih besar tapi tidak terlalu berlebihan
    height=500,  # Tinggi diperbesar agar tidak terlalu gepeng
    margin=dict(l=250, r=50, t=80, b=50),  # Lebarkan margin kiri agar label tidak terpotong
    font=dict(size=14),
    xaxis_title="Frequency Claim",
    yaxis_title="Claimant",
    yaxis=dict(autorange="reversed"),
    showlegend=False
)

    # Menampilkan angka di bar
    fig1.update_traces(textposition="auto")

    # Tampilkan di Streamlit
    st.plotly_chart(fig1)


Bulan_sev = df5.groupby("Month")["Gross Claim (In Million)"].sum().nlargest(12).reset_index()
Bulan_sev.columns=["Month", "Severity"]

with col2:
    st.markdown(
        """
        <style>
            .custom-title {
                text-align: center;
                margin-bottom: -20px;
                position: relative;
                top: 10px;
            }
        </style>
        <h4 class="custom-title">Top 10 Severity by Claimant</h4>
        """,
        unsafe_allow_html=True
    )

    # Horizontal Bar Plot untuk Frequency
    colors = [
    "#004d00",  
    "#006600",  
    "#008000",  
    "#009900",  
    "#00b300",  
    "#00cc00",  
    "#33ff33",  
    "#80ff80",  
    "#ccffcc",  
    "#ffffff"   
]
    fig2 = px.bar(Claimant_sev, 
                  x="Severity", 
                  y="Claimant", 
                  orientation="h", 
                  text="Severity", 
                  color="Claimant", 
                  color_discrete_sequence=colors)

    # Atur layout untuk memperbaiki tampilan
    fig2.update_layout(
    width=900,  # Lebar lebih besar tapi tidak terlalu berlebihan
    height=500,  # Tinggi diperbesar agar tidak terlalu gepeng
    margin=dict(l=250, r=50, t=80, b=50),  # Lebarkan margin kiri agar label tidak terpotong
    font=dict(size=14),
    xaxis_title="Severity",
    yaxis_title="Claimant",
    yaxis=dict(autorange="reversed"),
    showlegend=False
)

    # Menampilkan angka di bar
    fig2.update_traces(textposition="auto")

    # Tampilkan di Streamlit
    st.plotly_chart(fig2)

#Line Chart
# Pastikan df5 sudah ada
df5['Gross Claim (In Million)'] = pd.to_numeric(df5['Gross Claim (In Million)'], errors='coerce')

# Mapping angka bulan ke nama bulan
import streamlit as st
import pandas as pd
import plotly.express as px
import io  # Untuk menyimpan file sementara di memori

# Mapping angka ke nama bulan
month_mapping = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}
df5['Month'] = df5['Month'].replace(month_mapping)

# Agregasi data per bulan
monthly_claims = df5.groupby('Month', as_index=False)['Gross Claim (In Million)'].sum()

# Urutkan sesuai bulan (bukan alfabet)
order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
monthly_claims["Month"] = pd.Categorical(monthly_claims["Month"], categories=order, ordered=True)
monthly_claims = monthly_claims.sort_values("Month")

# **Bikin Judul Dinamis untuk Grafik & Excel**
title_parts = ["Gross Claims Per Month"]
filters_applied = []

if cob:
    filters_applied.append(f"Classification: {', '.join(cob)}")
if jenis:
    filters_applied.append(f" {', '.join(jenis)}")
if bisnis:
    filters_applied.append(f"Business: {', '.join(bisnis)}")
if bulan:
    filters_applied.append(f"Months: {', '.join(bulan)}")

if filters_applied:
    title_parts.append(" - ".join(filters_applied))

dynamic_title = " - ".join(title_parts)
filter_text = " | ".join(filters_applied) if filters_applied else "No filters applied"

# **Buat Line Chart dengan Judul Dinamis**
fig = px.line(
    monthly_claims, 
    x='Month', 
    y='Gross Claim (In Million)', 
    markers=True, 
    title=dynamic_title
)

fig.update_traces(line=dict(color="#D67230"))

# **Tampilkan Grafik di Streamlit**
st.plotly_chart(fig, use_container_width=True)

# **Simpan ke Excel dengan Keterangan**
output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Gross Claims")
    writer.sheets["Gross Claims"] = worksheet

    # **Tambahin Judul & Keterangan**
    worksheet.write("A1", "Data Gross OS Claim (In Million)")  
    worksheet.write("A3", "Filters Applied:")  # **Label Filter**
    worksheet.write("B3", filter_text)  # **Isi Filter**

    # **Tulis Data mulai dari baris ke-5**
    monthly_claims.to_excel(writer, sheet_name="Gross Claims", startrow=4, index=False)

# **Tombol Download**
st.download_button(
    label="Download Excel",
    data=output.getvalue(),
    file_name="Gross_Claims_Per_Month.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)