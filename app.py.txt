import streamlit as st
import pandas as pd
import math

# -------------------------------------------------------------------------
# [1] USPK 마스터 데이터 정의 (업로드된 최신 엑셀 시트 기준 고도화)
# -------------------------------------------------------------------------

# 1. 몰드별 Softgel encap 생산성 및 충진 기준
MOLD_DATA = {
    "3 Oval": {"base_mg": 150, "total_qty": 714, "rpm": 2.8, "pallet_caps": 2272000},
    "4 Oval": {"base_mg": 200, "total_qty": 585, "rpm": 2.8, "pallet_caps": 2000000},
    "5 Oval": {"base_mg": 275, "total_qty": 518, "rpm": 2.8, "pallet_caps": 1800000},
    "6 Oval": {"base_mg": 300, "total_qty": 468, "rpm": 2.8, "pallet_caps": 1600000},
    "7.5 Oval": {"base_mg": 375, "total_qty": 442, "rpm": 2.8, "pallet_caps": 1400000},
    "10 Oval": {"base_mg": 526, "total_qty": 396, "rpm": 2.8, "pallet_caps": 1200000},
    "16 Oblong": {"base_mg": 800, "total_qty": 272, "rpm": 2.8, "pallet_caps": 900000},
    "20 Oblong": {"base_mg": 1000, "total_qty": 224, "rpm": 3.0, "pallet_caps": 700000},
    "24 Oblong": {"base_mg": 1300, "total_qty": 210, "rpm": 3.0, "pallet_caps": 500000}
}

# 2. 피막 종류별 생산시간 보정 가중치 및 기본 단가
COATING_DATA = {
    "Clear": {"weight": 1.0, "base_cost_per_kg": 8500},
    "Color": {"weight": 1.15, "base_cost_per_kg": 9000},
    "Enteric (장용성 피막)": {"weight": 3.0, "base_cost_per_kg": 13000},
    "Enteric(Coating) (장용코팅 공정추가)": {"weight": 4.0, "base_cost_per_kg": 17000},
    "Veggie (식물성 캡슐)": {"weight": 2.2, "base_cost_per_kg": 22000}
}

# 3. 포장 형태별 Shift당 표준 작업량 및 배치 인원 (최신 시급 데이터 매칭)
PACKAGING_DATA = {
    "Bulk (벌크 포장)": {"output_per_shift": 500000, "workers": 3},
    "Bottle (일반 병포장)": {"output_per_shift": 5250, "workers": 5},
    "CSP Bottle": {"output_per_shift": 8000, "workers": 4},
    "Blister (PTP 포장)": {"output_per_shift": 120000, "workers": 4}
}

# 4. 시급 시트 기준 직급별 평균 시급 마스터
WAGE_DATA = {
    "최저시급/도급직": 10500,
    "사원급(고졸)": 14436,
    "사원급(대졸)": 15883,
    "주임급": 17996,
    "대리급": 21020,
    "과장급": 26128
}

# -------------------------------------------------------------------------
# [2] 웹 애플리케이션 UI 구성 (Streamlit)
# -------------------------------------------------------------------------
st.set_page_config(page_title="USPK 연질캡슐 통합 견적 시스템", layout="wide")

# 사이드바 로고 및 안내
st.sidebar.markdown("## ⚙️ USPK ERP 시스템")
st.sidebar.info("본 시스템은 유에스파마텍코리아(주)의 영업 보안 자산입니다. 외부 무단 유출을 금지합니다.")

st.title("📊 연질캡슐(Softgel) 내부 원가 및 견적 시뮬레이터")
st.markdown("사내 직원 전용 웹 도구입니다. 원료비, 공정 변수 선택 시 실시간으로 내부 원가와 DDP/EXW 단가가 자동 산출됩니다.")
st.divider()

# 레이아웃 분할
col_input, col_result = st.columns([1, 1.4])

with col_input:
    st.subheader("📋 기본 정보 및 공정 조건")
    
    with st.expander("1. 프로젝트 기본 정보", expanded=True):
        client_name = st.text_input("고객사명 (Client)", "NOW")
        product_name = st.text_input("제품명 (Product Name)", "DHA-250 mg Fish Oil Softgel (51610E)")
        order_qty = st.number_input("발주 수량 (총 알약 수량, ea)", min_value=1000, value=1000000, step=100000)
        pack_in_qty = st.number_input("포장 단위당 입수량 (예: 병당 알약수 / Bulk는 그대로 입력)", min_value=1, value=10000)

    with st.expander("2. 제형 및 생산 설비 매칭", expanded=True):
        capsule_size = st.selectbox("캡슐 사이즈 (Mold Size)", list(MOLD_DATA.keys()), index=5) # Default: 10 Oval
        coating_type = st.selectbox("피막 종류 (Shell Type)", list(COATING_DATA.keys()), index=0) # Default: Clear
        packaging_type = st.selectbox("포장 형태 (Packaging Type)", list(PACKAGING_DATA.keys()), index=0) # Default: Bulk
        
    with st.expander("3. 노무비 및 원료 원가 변수", expanded=True):
        selected_wage_group = st.selectbox("포장/성형 적용 직급 시급", list(WAGE_DATA.keys()), index=2)
        hourly_wage = WAGE_DATA[selected_wage_group]
        
        main_ingredient_cost = st.number_input("캡슐 1알당 내용물(주약) 원가 (원)", min_value=0.0, value=28.5, step=1.0)
        loss_rate_input = st.slider("제조 수율 손실률 (Loss %)", min_value=1.0, max_value=15.0, value=5.0, step=0.5) / 100
        target_profit_rate = st.slider("목표 이익률 (%)", min_value=5, max_value=60, value=25, step=5) / 100

# -------------------------------------------------------------------------
# [3] 내부 견적 연산 엔진 (Excel 원가 분석 로직 완벽 복제)
# -------------------------------------------------------------------------
# 1. 이론 및 실제 생산 수량 계산 (로스분 반영)
total_required_capsules = order_qty * (1 + loss_rate_input)

# 2. 성형(Incap) 생산성 및 노무비 산출
mold = MOLD_DATA[capsule_size]
coating = COATING_DATA[coating_type]

theoretical_hourly_output = mold["total_qty"] * mold["rpm"] * 60
# 피막 가중치 적용한 실제 시간당 생산량 계산
actual_hourly_output = theoretical_hourly_output / coating["weight"]
required_incap_hours = total_required_capsules / actual_hourly_output
# 성형 공정 기본 배치 인원 (조장1, 조원2 기준 총 3명 적용)
incap_labor_cost = required_incap_hours * 3 * hourly_wage

# 3. 젤라틴 매스(Shell) 비용 계산
# 충진량 비례 피막 젤라틴 소요 중량 예측 연산 (톤당 단가 환산)
gelatin_mass_per_capsule_kg = (mold["base_mg"] * 0.45) / 1000000
total_gelatin_cost = total_required_capsules * gelatin_mass_per_capsule_kg * coating["base_cost_per_kg"]

# 4. 패키징 원가 계산
pkg = PACKAGING_DATA[packaging_type]
total_packing_units = order_qty / pack_in_qty
pkg_hourly_output = pkg["output_per_shift"] / 8
required_pkg_hours = total_packing_units / pkg_hourly_output
pkg_labor_cost = required_pkg_hours * pkg["workers"] * hourly_wage

# 5. 종합 원가 및 최종 제안가 빌드업 (판관비 11.025% 반영 및 환율 적용)
total_material_cost = (order_qty * main_ingredient_cost) + total_gelatin_cost
total_manufacturing_cost = total_material_cost + incap_labor_cost + pkg_labor_cost
sg_and_a_cost = total_manufacturing_cost * 0.11025  # 최신 USPK 엑셀의 판관비율 고정값
grand_total_cost = total_manufacturing_cost + sg_and_a_cost

# 마진율을 포함한 최종 매출 총액
final_revenue = grand_total_cost / (1 - target_profit_rate)
cost_per_1000_caps = (final_revenue / order_qty) * 1000

# 물류 지표 계산 (Pallet 적재 시뮬레이션)
required_pallets = math.ceil(order_qty / mold["pallet_caps"])

# -------------------------------------------------------------------------
# [4] 우측 결과 대시보드 및 프로포마 인보이스 출력
# -------------------------------------------------------------------------
with col_result:
    st.subheader("📊 원가 대시보드 & 시뮬레이션 리포트")
    
    # 주요 메트릭 지표 상단 배치
    m1, m2, m3 = st.columns(3)
    m2.metric("최종 견적가 (총액)", f"{int(final_revenue):,} 원")
    m1.metric("1,000 캡슐당 단가", f"{round(cost_per_1000_caps, 2):,} 원")
    m3.metric("필요 물류 팔레트(PLT)", f"{required_pallets} PLT")
    
    st.divider()
    
    # 상세 탭 구성
    tab_invoice, tab_breakdown = st.tabs(["📄 Proforma Invoice 양식", "🔍 제조원가 상세 분석"])
    
    with tab_invoice:
        st.markdown(f"### PROFORMA INVOICE (견적서)")
        st.caption("유에스파마텍코리아 주식회사 | 충북 음성군 원남산단로 111")
        
        invoice_html = f"""
        <div style="border:1px solid #d3d3d3; padding:15px; border-radius:5px; background-color:#fafafa; font-family:sans-serif;">
            <table style="width:100%; border-collapse: collapse;">
                <tr><td><b>Customer:</b> {client_name}</td><td style="text-align:right;"><b>Date:</b> 2026-05-21</td></tr>
                <tr><td><b>Product:</b> {product_name}</td><td style="text-align:right;"><b>Dosage Form:</b> Softgel ({capsule_size})</td></tr>
                <tr><td><b>Order Qty:</b> {order_qty:,} ea</td><td style="text-align:right;"><b>Shell Type:</b> {coating_type}</td></tr>
            </table>
            <hr style="border:0.5px solid #d3d3d3;">
            <table style="width:100%; margin-top:10px;">
                <tr style="background-color:#e6e6e6; font-weight:bold;">
                    <td style="padding:8px;">Description</td>
                    <td style="text-align:right; padding:8px;">Quantity</td>
                    <td style="text-align:right; padding:8px;">Unit Price (per 1,000sg)</td>
                    <td style="text-align:right; padding:8px;">Amount</td>
                </tr>
                <tr>
                    <td style="padding:8px;">{product_name}<br><small style="color:gray;">Packaging: {packaging_type}</small></td>
                    <td style="text-align:right; padding:8px;">{order_qty:,} ea</td>
                    <td style="text-align:right; padding:8px;">₩ {round(cost_per_1000_caps, 2):,}</td>
                    <td style="text-align:right; padding:8px;"><b>₩ {int(final_revenue):,}</b></td>
                </tr>
            </table>
        </div>
        """
        st.markdown(invoice_html, unsafe_style_allowed=True)
        
    with tab_breakdown:
        st.markdown("#### 💡 항목별 제조 원가 명세")
        
        breakdown_data = {
            "원가 구성 항목": [
                "1. 원자재비 (주약 내용물 비용)", 
                "2. 피막 재료비 (젤라틴 매스)", 
                "3. 성형(Incap) 라인 노무비", 
                "4. 패키징 공정 노무비", 
                "5. 일반관리비/판관비 (11.025%)",
                "6. 내부 산출 총원가",
                "7. 영업 마진액 (설정 범위 내)"
            ],
            "금액 (₩)": [
                int(order_qty * main_ingredient_cost),
                int(total_gelatin_cost),
                int(incap_labor_cost),
                int(pkg_labor_cost),
                int(sg_and_a_cost),
                int(grand_total_cost),
                int(final_revenue - grand_total_cost)
            ]
        }
        df_breakdown = pd.DataFrame(breakdown_data)
        st.table(df_breakdown.set_index("원가 구성 항목"))
        
        # 원가 비율 그래프 표시
        st.bar_chart(df_breakdown.iloc[:5].set_index("원가 구성 항목")["금액 (₩)"])
