import streamlit as st
import pandas as pd

st.set_page_config(page_title="Assignment Portal", page_icon="🎯", layout="wide")

st.markdown("""
<style>
    .main-header {font-size: 2.5rem; font-weight: 700; color: #1f77b4; text-align: center; margin-bottom: 1rem;}
    .upload-section {background: #f0f2f6; padding: 2rem; border-radius: 10px; margin: 2rem 0;}
    .metric-card {background: white; padding: 1.5rem; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    .stButton>button {width: 100%; font-size: 1.1rem; font-weight: 600; padding: 0.8rem;}
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">🎯 Faculty Assignment Portal</p>', unsafe_allow_html=True)

st.markdown('<div class="upload-section">', unsafe_allow_html=True)
uploaded = st.file_uploader("Upload CSV File", type=['csv'], label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

if uploaded:
    df = pd.read_csv(uploaded)
    
    if 'processed' not in st.session_state:
        st.session_state.processed = False
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if st.button("🚀 Execute Assignment", type="primary"):
            with st.spinner("Processing..."):
                # Extract faculty identifiers
                fac_list = [c for c in df.columns if c not in ['Roll', 'Name', 'Email', 'CGPA']]
                fac_count = len(fac_list)
                
                # Order by academic performance
                ordered_data = df.sort_values(by='CGPA', ascending=False).reset_index(drop=True)
                
                # Storage structures
                results = []
                pref_matrix = {f: [0] * 18 for f in fac_list}
                
                total_count = len(ordered_data)
                idx_position = 0
                
                # Process in cohorts
                while idx_position < total_count:
                    cohort_limit = min(fac_count, total_count - idx_position)
                    cohort_end = idx_position + cohort_limit
                    
                    current_cohort = ordered_data.iloc[idx_position:cohort_end]
                    used_facs = set()
                    
                    for _, record in current_cohort.iterrows():
                        # Build choice map
                        choice_map = []
                        for f in fac_list:
                            choice_map.append((int(record[f]), f))
                        
                        choice_map.sort()
                        
                        # Find available faculty
                        selected_fac = None
                        selected_rank = None
                        
                        for rank, f in choice_map:
                            if f not in used_facs:
                                selected_fac = f
                                selected_rank = rank
                                used_facs.add(f)
                                break
                        
                        # Fallback mechanism
                        if selected_fac is None:
                            for f in fac_list:
                                if f not in used_facs:
                                    selected_fac = f
                                    selected_rank = int(record[f])
                                    used_facs.add(f)
                                    break
                        
                        if selected_fac:
                            results.append({
                                'Roll': record['Roll'],
                                'Name': record['Name'],
                                'Email': record['Email'],
                                'CGPA': record['CGPA'],
                                'Allocated': selected_fac,
                                'Preference_Rank': selected_rank
                            })
                            
                            pref_matrix[selected_fac][selected_rank - 1] += 1
                    
                    idx_position = cohort_end
                
                # Build output structures
                assignment_data = pd.DataFrame(results)
                
                # Statistics table
                stat_records = []
                for f in fac_list:
                    stat_entry = {'Fac': f}
                    for i in range(18):
                        stat_entry[f'Count Pref {i+1}'] = pref_matrix[f][i]
                    stat_records.append(stat_entry)
                
                stats_data = pd.DataFrame(stat_records)
                summary_data = assignment_data.groupby('Allocated').size().reset_index(name='Student_Count')
                
                st.session_state.assignment_data = assignment_data
                st.session_state.stats_data = stats_data
                st.session_state.summary_data = summary_data
                st.session_state.processed = True
    
    if st.session_state.processed:
        st.success("✅ Assignment Complete")
        
        # Metrics
        st.markdown("### 📊 Summary Metrics")
        m1, m2, m3, m4 = st.columns(4)
        
        with m1:
            st.metric("Students", len(st.session_state.assignment_data))
        with m2:
            st.metric("Faculties", len(st.session_state.stats_data))
        with m3:
            avg_val = len(st.session_state.assignment_data) / len(st.session_state.stats_data)
            st.metric("Avg/Faculty", f"{avg_val:.1f}")
        with m4:
            avg_pref = st.session_state.assignment_data['Preference_Rank'].mean()
            st.metric("Avg Rank", f"{avg_pref:.2f}")
        
        # Statistics
        st.markdown("### 📈 Faculty Distribution")
        st.dataframe(st.session_state.summary_data, use_container_width=True, hide_index=True)
        st.bar_chart(st.session_state.summary_data.set_index('Allocated')['Student_Count'])
        
        st.markdown("### 📋 Preference Statistics")
        st.dataframe(st.session_state.stats_data, use_container_width=True, hide_index=True)
        
        # Downloads
        st.markdown("### 💾 Export Results")
        dl1, dl2 = st.columns(2)
        
        with dl1:
            output_main = st.session_state.assignment_data[['Roll', 'Name', 'Email', 'CGPA', 'Allocated']].copy()
            csv_main = output_main.to_csv(index=False)
            st.download_button(
                "📥 Download Assignments",
                csv_main,
                "output_btp_mtp_allocation.csv",
                "text/csv",
                use_container_width=True
            )
        
        with dl2:
            csv_stats = st.session_state.stats_data.to_csv(index=False)
            st.download_button(
                "📥 Download Statistics",
                csv_stats,
                "fac_preference_count.csv",
                "text/csv",
                use_container_width=True
            )
else:
    st.info("📁 Upload a CSV file to start")