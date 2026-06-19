import numpy as np
import pandas as pd

def generate_sustainability_panel_data(seed=42):
    np.random.seed(seed)
    
    companies = [f"C{i:03d}" for i in range(1, 21)]
    company_names = [
        "Aksa Enerji", "Ford Otosan", "Akbank", "Arçelik", "Tüpraş",
        "Aselsan", "BİM Mağazalar", "Koç Holding", "Sabancı Holding", "Migros",
        "Şişecam", "Türk Hava Yolları", "Garanti BBVA", "Ereğli Demir Çelik", "Petkim",
        "Kardemir", "Mavi Giyim", "Tofaş", "Ülker", "Pegasus"
    ]
    
    sectors = [
        "Energy", "Automotive", "Banking", "Consumer Durables", "Energy / Refining",
        "Defense", "Retail Trade", "Holding", "Holding", "Retail Trade",
        "Glass / Industrial", "Aviation", "Banking", "Steel / Heavy", "Petrochemicals",
        "Steel / Heavy", "Textile / Retail", "Automotive", "Food / Consumer", "Aviation"
    ]
    
    years = [2020, 2021, 2022, 2023, 2024]
    
    rows = []
    
    # Baseline stats for each company to maintain panel structure (firm fixed effects)
    firm_baselines = {}
    for comp, name, sect in zip(companies, company_names, sectors):
        firm_baselines[comp] = {
            "name": name,
            "sector": sect,
            # Environmental raw baselines
            "scope1": np.random.uniform(100000, 2000000),
            "scope2": np.random.uniform(20000, 500000),
            "scope3": np.random.uniform(500000, 10000000),
            "renewable_ratio": np.random.uniform(5.0, 40.0),
            "energy_consumption": np.random.uniform(1000000, 30000000),
            "water_consumption": np.random.uniform(500000, 6000000),
            "waste_ton": np.random.uniform(2000, 80000),
            "net_zero_prob": np.random.uniform(0.1, 0.7),
            "reduction_prob": np.random.uniform(0.2, 0.8),
            
            # Social baselines
            "female_emp_ratio": np.random.uniform(15.0, 45.0),
            "training_hours": np.random.uniform(15.0, 50.0),
            "accident_rate": np.random.uniform(0.2, 3.5),
            
            # Governance baselines
            "board_ind_ratio": np.random.uniform(25.0, 50.0),
            "female_board_ratio": np.random.uniform(10.0, 30.0),
            "esg_comm_prob": np.random.uniform(0.3, 0.9),
            "ethics_prob": np.random.uniform(0.5, 0.95),
            "anti_corr_prob": np.random.uniform(0.5, 0.95),
            
            # Text / LLM baselines
            "gri_score": np.random.uniform(50.0, 85.0),
            "tcfd_score": np.random.uniform(40.0, 75.0),
            "vague_ratio": np.random.uniform(35.0, 65.0),
            "quant_evidence": np.random.uniform(20.0, 55.0),
            "evidence_backed": np.random.uniform(25.0, 60.0),
            "sentiment_score": np.random.uniform(45.0, 75.0),
            
            # Financial baselines
            "size": np.random.uniform(10.0, 13.5),
            "roa": np.random.uniform(2.0, 10.0),
            "leverage": np.random.uniform(30.0, 75.0)
        }
    
    for comp in companies:
        base = firm_baselines[comp]
        for year_idx, year in enumerate(years):
            # Time trends
            trend = year_idx * 0.05  # positive trend (e.g. reduction in emissions, rise in scores)
            
            # 1. Environmental Variables (with a downward trend for emissions and upward for renewable)
            s1_decay = max(0.6, 1.0 - trend * np.random.uniform(0.5, 1.2))
            scope1 = base["scope1"] * s1_decay * np.random.uniform(0.95, 1.05)
            scope2 = base["scope2"] * s1_decay * np.random.uniform(0.95, 1.05)
            scope3 = base["scope3"] * s1_decay * np.random.uniform(0.95, 1.05)
            total_ghg = scope1 + scope2 + scope3
            
            ren_growth = 1.0 + trend * np.random.uniform(0.8, 1.5)
            renewable_energy_ratio = min(100.0, base["renewable_ratio"] * ren_growth * np.random.uniform(0.98, 1.02))
            
            energy_efficiency = max(0.8, 1.0 - trend * np.random.uniform(0.2, 0.6))
            energy_consumption_mwh = base["energy_consumption"] * energy_efficiency * np.random.uniform(0.97, 1.03)
            water_consumption_m3 = base["water_consumption"] * energy_efficiency * np.random.uniform(0.97, 1.03)
            waste_ton = base["waste_ton"] * energy_efficiency * np.random.uniform(0.97, 1.03)
            
            # Binary variables (higher probability in later years)
            net_zero_target = 1 if np.random.rand() < min(0.95, base["net_zero_prob"] * (1.0 + trend * 1.5)) else 0
            emission_reduction_target = 1 if np.random.rand() < min(0.95, base["reduction_prob"] * (1.0 + trend * 1.2)) else 0
            
            target_year = np.random.choice([2030, 2035, 2040, 2050]) if net_zero_target == 1 else np.nan
            
            # 2. Social Variables
            soc_growth = 1.0 + trend * np.random.uniform(0.3, 0.8)
            female_employee_ratio = min(100.0, base["female_emp_ratio"] * soc_growth * np.random.uniform(0.98, 1.02))
            training_hours_per_employee = base["training_hours"] * soc_growth * np.random.uniform(0.95, 1.05)
            
            accident_decay = max(0.4, 1.0 - trend * np.random.uniform(0.5, 1.2))
            occupational_accident_rate = base["accident_rate"] * accident_decay * np.random.uniform(0.90, 1.10)
            
            # 3. Governance Variables
            board_independence_ratio = min(100.0, base["board_ind_ratio"] * (1.0 + trend * 0.2) * np.random.uniform(0.98, 1.02))
            female_board_ratio = min(100.0, base["female_board_ratio"] * (1.0 + trend * 0.4) * np.random.uniform(0.98, 1.02))
            
            esg_committee = 1 if np.random.rand() < min(0.98, base["esg_comm_prob"] * (1.0 + trend * 1.0)) else 0
            ethics_policy = 1 if np.random.rand() < min(0.99, base["ethics_prob"] * (1.0 + trend * 0.5)) else 0
            anti_corruption_policy = 1 if np.random.rand() < min(0.99, base["anti_corr_prob"] * (1.0 + trend * 0.5)) else 0
            
            # 4. Text-based Variables (increasing quality over time)
            text_trend = trend * np.random.uniform(0.5, 1.2)
            gri_score = min(100.0, base["gri_score"] + text_trend * 15.0 + np.random.uniform(-2, 2))
            tcfd_score = min(100.0, base["tcfd_score"] + text_trend * 20.0 + np.random.uniform(-2, 2))
            
            vague_statement_ratio = max(5.0, base["vague_ratio"] - text_trend * 25.0 + np.random.uniform(-3, 3))
            quantitative_evidence_ratio = min(100.0, base["quant_evidence"] + text_trend * 20.0 + np.random.uniform(-3, 3))
            evidence_backed_claim_ratio = min(100.0, base["evidence_backed"] + text_trend * 22.0 + np.random.uniform(-3, 3))
            
            # Positive sentiment (should be relatively stable, maybe slight rise)
            positive_sentiment_score = min(100.0, base["sentiment_score"] + np.random.uniform(-5, 5))
            
            # 5. Financial Control Variables
            firm_size_log_assets = base["size"] * (1.0 + year_idx * 0.01) + np.random.uniform(-0.05, 0.05)
            roa = base["roa"] + np.random.uniform(-1.5, 1.5)
            leverage = base["leverage"] + np.random.uniform(-2.0, 2.0)
            
            rows.append({
                "company_id": comp,
                "company_name": base["name"],
                "year": year,
                "sector": base["sector"],
                "country": "Turkey",
                "report_url": f"https://www.example.com/{base['name'].lower().replace(' ', '_')}_{year}_report.pdf",
                
                # Environmental (ENV) Raw
                "scope1_tco2e": scope1,
                "scope2_tco2e": scope2,
                "scope3_tco2e": scope3,
                "total_ghg_tco2e": total_ghg,
                "renewable_energy_ratio": renewable_energy_ratio,
                "energy_consumption_mwh": energy_consumption_mwh,
                "water_consumption_m3": water_consumption_m3,
                "waste_ton": waste_ton,
                "net_zero_target": net_zero_target,
                "emission_reduction_target": emission_reduction_target,
                "target_year": target_year,
                
                # Social (SOC) Raw
                "female_employee_ratio": female_employee_ratio,
                "training_hours_per_employee": training_hours_per_employee,
                "occupational_accident_rate": occupational_accident_rate,
                
                # Governance (GOV) Raw
                "board_independence_ratio": board_independence_ratio,
                "female_board_ratio": female_board_ratio,
                "esg_committee": esg_committee,
                "ethics_policy": ethics_policy,
                "anti_corruption_policy": anti_corruption_policy,
                
                # Text Raw
                "gri_score": gri_score,
                "tcfd_score": tcfd_score,
                "vague_statement_ratio": vague_statement_ratio,
                "quantitative_evidence_ratio": quantitative_evidence_ratio,
                "evidence_backed_claim_ratio": evidence_backed_claim_ratio,
                "positive_sentiment_score": positive_sentiment_score,
                
                # Financial Control Raw
                "firm_size_log_assets": firm_size_log_assets,
                "roa": roa,
                "leverage": leverage
            })
            
    df = pd.DataFrame(rows)
    return df

if __name__ == "__main__":
    df = generate_sustainability_panel_data()
    df.to_csv("sustainability_panel_raw_data.csv", index=False)
    print("Generated data shape:", df.shape)
    print("Unique companies:", len(df["company_id"].unique()))
    print("Years:", df["year"].unique().tolist())
