import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Smart Mapping Engine (Fuzzy similarity)
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# ----------------------------------------------------
# 1. Utilities for Validation & Intelligent Mapping
# ----------------------------------------------------

def is_valid_email(email):
    """Checks email validity via regex."""
    if not email or pd.isna(email): return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """
    Intelligently maps names (e.g., 'John Doe Manager' to 'John Doe') 
    using similarity analysis.
    """
    target_name = str(target_name).strip()
    if not target_name or target_name == "nan": return None, "Name Empty"
    
    # Check exact match first
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    # Similarity analysis (Threshold: 90%)
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 90:
            suggested_name = result[0]
            return contact_dict[suggested_name], f"Suggested({suggested_name})"
    
    return None, "Not Found"

# ----------------------------------------------------
# 2. Tabs: Automated Notification & Validation
# ----------------------------------------------------

# (Assuming tab_alert is defined in your st.tabs list)
with tab_alert:
    st.subheader("📨 Intelligent Manager Notifications & Verification")
    
    if 'manager_contacts' not in locals() or not manager_contacts:
        st.warning("⚠️ Contact mapping file is required.")
    else:
        # Filter for high-risk unmatched contracts
        targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        if targets.empty:
            st.success("🎉 No high-risk unmatched targets found.")
        else:
            st.info("🔍 Performing manager mapping and email verification...")

            # Verification logic replacing the SyntaxError line
            verify_list = []
            for _, row in targets.iterrows():
                mgr_name = row["구역담당자_통합"]
                info, status = get_smart_contact(mgr_name, manager_contacts)
                email = info.get("email", "") if info else ""
                
                verify_list.append({
                    "Contract_ID": row["계약번호_정제"],
                    "Client": row["상호"],
                    "Manager": mgr_name,
                    "Mapped_Email": email,
                    "Mapping_Status": status,
                    "Address_Valid": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # Summary Metrics for Data Integrity
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.metric("Mapping Accuracy", f"{(v_df['Mapping_Status'] != 'Not Found').mean()*100:.1f}%")
            with col_v2:
                invalid_cnt = v_df[~v_df["Address_Valid"] & (v_df["Mapped_Email"] != "")].shape[0]
                st.metric("Address Errors", f"{invalid_cnt} items", delta_color="inverse")
            with col_v3:
                st.metric("Target Contracts", f"{len(v_df)} units")

            st.markdown("---")

            # Notification Editor
            st.markdown("#### 🛠️ Final Review of Mailing List")
            agg_targets = v_df.groupby(["Manager", "Mapped_Email", "Mapping_Status", "Address_Valid"]).size().reset_index(name="Count")
            
            edited_agg = st.data_editor(
                agg_targets,
                column_config={
                    "Mapped_Email": st.column_config.TextColumn("Email (Editable)", required=True),
                    "Count": st.column_config.NumberColumn("Contracts", disabled=True),
                    "Address_Valid": st.column_config.CheckboxColumn("Valid Format", disabled=True)
                },
                use_container_width=True,
                key="alert_editor_pro",
                hide_index=True
            )

            # Execution logic (SMTP send omitted for brevity - follow existing patterns)
