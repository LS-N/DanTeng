@staticmethod
    def render_native_editor(desc, subset, is_edit, cols_a, cols_b):
        st.markdown(f'<div class="info-bar">ℹ️ {desc}</div>', unsafe_allow_html=True)
        
        # 1. 准备数据
        df_display = subset[['序号', '目标字段', '源表', '匹配字段', '逻辑说明']].copy().reset_index(drop=True)
        
        # 【关键技巧】将序号转换为字符串，这样 Streamlit 就会默认将其“居左对齐”，
        # 从而在视觉上与其他列保持一致。
        df_display['序号'] = df_display['序号'].astype(str)
        
        # 2. 列配置 (设置宽度 width 和类型)
        column_config = {
            # width="small": 让序号列尽可能窄，不占用多余空间
            "序号": st.column_config.TextColumn("序号", width="small", disabled=True),
            
            # width="medium": 默认宽度
            "目标字段": st.column_config.TextColumn("目标字段", disabled=True, width="medium"),
            
            # width="large": 给逻辑说明更多空间，避免换行太多
            "逻辑说明": st.column_config.TextColumn("逻辑说明", disabled=True, width="large"),
        }

        # 3. 动态配置可编辑列
        if is_edit:
            column_config["源表"] = st.column_config.SelectboxColumn(
                "源表", options=["Source A", "Source B"], width="small", required=True
            )
            column_config["匹配字段"] = st.column_config.SelectboxColumn(
                "匹配字段", options=cols_a + cols_b, width="medium", required=True
            )
        else:
            column_config["源表"] = st.column_config.TextColumn("源表", disabled=True)
            column_config["匹配字段"] = st.column_config.TextColumn("匹配字段", disabled=True)

        # 4. 动态计算表格高度
        calc_height = (len(df_display) + 1) * 35 + 10
        final_height = max(400, min(1000, calc_height))

        # 5. 渲染表格
        edited = st.data_editor(
            df_display,
            column_config=column_config,
            use_container_width=True, # 这让表格占满容器宽度
            hide_index=True,          # 隐藏默认的索引列
            disabled=not is_edit,
            height=final_height,
            key=f"editor_{subset.iloc[0]['所属表']}"
        )

        # 6. 保存逻辑 (保持不变)
        if is_edit:
            for i, row in edited.iterrows():
                orig_idx = subset.index[i]
                orig_row = st.session_state.mapping_config.loc[orig_idx]
                
                if 'Source' not in orig_row['源表']:
                    continue 
                
                if row['源表'] != orig_row['源表']:
                    st.session_state.mapping_config.at[orig_idx, '源表'] = row['源表']
                    target_opts = cols_a if row['源表'] == 'Source A' else cols_b
                    new_val = row['目标字段'] if row['目标字段'] in target_opts else (target_opts[0] if target_opts else None)
                    st.session_state.mapping_config.at[orig_idx, '匹配字段'] = new_val
                elif row['匹配字段'] != orig_row['匹配字段']:
                    valid_opts = cols_a if row['源表'] == 'Source A' else cols_b
                    if row['匹配字段'] in valid_opts:
                        st.session_state.mapping_config.at[orig_idx, '匹配字段'] = row['匹配字段']
