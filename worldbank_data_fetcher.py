import os
import time
import requests
import pandas as pd
import numpy as np

def fetch_worldbank_indicator(countries_str, indicator_code, start_year, end_year):
    """Fetch a single indicator from World Bank API."""
    url = (
        f"https://api.worldbank.org/v2/country/{countries_str}/indicator/{indicator_code}"
        f"?format=json&per_page=20000&date={start_year}:{end_year}"
    )
    
    # Try with retries
    for attempt in range(3):
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            payload = response.json()
            
            if len(payload) < 2 or payload[1] is None:
                return pd.DataFrame()
                
            rows = []
            for item in payload[1]:
                val = item.get("value")
                rows.append({
                    "country": item.get("country", {}).get("value"),
                    "countryiso3code": item.get("countryiso3code"),
                    "year": int(item.get("date")),
                    "indicator_code": indicator_code,
                    "value": val if val is not None else np.nan
                })
            return pd.DataFrame(rows)
        except Exception as exc:
            print(f"Attempt {attempt+1} failed for {indicator_code}. Error: {exc}")
            time.sleep(1)
    return pd.DataFrame()

def download_and_clean_wb_data():
    countries = ["TUR", "DEU", "FRA", "ITA", "ESP", "GBR", "USA", "JPN", "KOR", "CAN", 
                 "AUS", "SWE", "NOR", "NLD", "CHE", "MEX", "BRA", "CHN", "IND", "ZAF"]
    countries_str = ";".join(countries)
    
    start_year = 2010
    end_year = 2023
    
    indicators = {
        # Environmental (15)
        "EN.GHG.CO2.PC.CE.AR5": "CO2_per_capita",
        "EN.GHG.CO2.MT.CE.AR5": "CO2_kt",
        "EN.GHG.ALL.MT.CE.AR5": "Total_GHG",
        "EN.GHG.CH4.MT.CE.AR5": "Methane_emissions",
        "EN.GHG.N2O.MT.CE.AR5": "Nitrous_oxide_emissions",
        "EG.FEC.RNEW.ZS": "Renewable_energy_consumption_pct",
        "EG.ELC.RNEW.ZS": "Renewable_electricity_pct",
        "AG.LND.FRST.ZS": "Forest_area_pct",
        "EN.ATM.PM25.MC.M3": "PM25_pollution",
        "EG.ELC.ACCS.ZS": "Access_to_electricity",
        "NY.GDP.TOTL.RT.ZS": "Natural_resources_rents_pct",
        "AG.LND.AGRI.ZS": "Agricultural_land_pct",
        "ER.PTD.TOTL.ZS": "Protected_areas_pct",
        "EG.USE.COMM.GD.PP.KD": "Energy_use_kg_oil_equiv",
        "EG.ELC.COAL.ZS": "Electricity_from_coal_pct",
        
        # Social (14)
        "SP.DYN.LE00.IN": "Life_expectancy",
        "SH.DYN.MORT": "Under_5_mortality_rate",
        "SH.H2O.SMDW.ZS": "Safe_drinking_water_pct",
        "SH.STA.SMSS.ZS": "Safe_sanitation_pct",
        "SE.PRM.ENRR": "Primary_school_enrollment",
        "SE.SEC.ENRR": "Secondary_school_enrollment",
        "SE.TER.ENRR": "Tertiary_school_enrollment",
        "SE.ADT.LITR.ZS": "Adult_literacy_rate",
        "SL.UEM.TOTL.ZS": "Unemployment_pct",
        "SL.TLF.CACT.FE.ZS": "Female_labor_participation_pct",
        "SI.POV.GINI": "Gini_index",
        "SI.POV.NAHC": "Poverty_headcount_ratio",
        "EG.CFT.ACCS.ZS": "Clean_fuel_access_pct",
        "IQ.CPA.GNDR.XQ": "CPIA_gender_equality_rating",
        
        # Governance / Economic (12)
        "NY.GDP.MKTP.KD.ZG": "GDP_growth_pct",
        "NY.GDP.PCAP.KD": "GDP_per_capita_constant_usd",
        "FP.CPI.TOTL.ZG": "Inflation_pct",
        "NV.IND.TOTL.ZS": "Industry_value_added_pct_GDP",
        "NV.SRV.TOTL.ZS": "Services_value_added_pct_GDP",
        "NV.AGR.TOTL.ZS": "Agriculture_value_added_pct_GDP",
        "BX.KLT.DINV.WD.GD.ZS": "FDI_net_inflow_pct_GDP",
        "GB.XPD.RSDV.GD.ZS": "RD_expenditure_pct_GDP",
        "TX.VAL.TECH.MF.ZS": "High_tech_exports_pct_manufactured",
        "IP.PAT.RESD": "Patent_applications_residents",
        "MS.MIL.XPND.GD.ZS": "Military_expenditure_pct_GDP",
        "SE.XPD.TOTL.GD.ZS": "Education_expenditure_pct_GDP"
    }
    
    all_dfs = []
    print(f"Starting WDI download for {len(indicators)} indicators across {len(countries)} countries (2010-2023)...")
    
    for i, (code, name) in enumerate(indicators.items(), 1):
        print(f"[{i}/{len(indicators)}] Fetching {code} ({name})...")
        df = fetch_worldbank_indicator(countries_str, code, start_year, end_year)
        if not df.empty:
            df["indicator_name"] = name
            all_dfs.append(df)
        time.sleep(0.5) # respect API rate limits
        
    if not all_dfs:
        print("Error: No data downloaded.")
        return
        
    # Combine long format data
    combined_long = pd.concat(all_dfs, ignore_index=True)
    
    # Pivot to wide format
    panel_df = combined_long.pivot_table(
        index=["country", "countryiso3code", "year"],
        columns="indicator_name",
        values="value",
        aggfunc="first"
    ).reset_index()
    
    # Sort
    panel_df = panel_df.sort_values(["countryiso3code", "year"]).reset_index(drop=True)
    
    # Cleaning missing values:
    # 1. Fill CPIA (which is only for low income, so it might be NaN for OECD, fill with 6.0/maximum or drop, or fill with median)
    # Let's fill NaNs per country using linear interpolation (helps with gaps in Gini, literacy etc.)
    print("Cleaning missing values via group interpolation...")
    panel_df = panel_df.groupby("country", group_keys=False).apply(
        lambda g: g.interpolate(method='linear', limit_direction='both')
    )
    
    # 2. For remaining NaNs (where country has NO data for that indicator at all), fill with column median
    for col in panel_df.columns:
        if col not in ["country", "countryiso3code", "year"] and panel_df[col].isna().any():
            median_val = panel_df[col].median()
            # If all are NaN for some reason, fill with 0 or a sensible default
            if pd.isna(median_val):
                median_val = 0.0
            panel_df[col] = panel_df[col].fillna(median_val)
            
    # Save raw data for analysis script
    # Map country codes to match company_id / company_name in analysis script
    # We rename 'country' -> 'company_name' and 'countryiso3code' -> 'company_id' 
    # to maintain compatibility with the core analysis script without rewriting too much code!
    panel_df = panel_df.rename(columns={
        "country": "company_name",
        "countryiso3code": "company_id"
    })
    
    # Save to CSV
    output_path = "sustainability_panel_raw_data.csv"
    panel_df.to_csv(output_path, index=False)
    print(f"Data saved to {output_path}. Shape: {panel_df.shape}")
    return panel_df

if __name__ == "__main__":
    download_and_clean_wb_data()
