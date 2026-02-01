import pandas as pd
import json
import ollama
import time
import re


# -------------------------------
# Utility Functions
# -------------------------------

def safe_join(values):
    """Safely join list values as comma-separated string"""
    if not values:
        return ""
    return ",".join(map(str, values))


def expand_service_ranges(service_codes):
    """
    Expand ranges like 99242-99245 into individual codes
    """
    expanded = []

    for code in service_codes:
        code_str = str(code)

        # Check for range pattern like 99242-99245
        match = re.match(r"(\d+)\s*-\s*(\d+)", code_str)
        if match:
            start = int(match.group(1))
            end = int(match.group(2))
            for num in range(start, end + 1):
                expanded.append(str(num))
        else:
            expanded.append(code_str)

    return expanded


# -------------------------------
# LLM Call Function
# -------------------------------

def call_llm(definition):

    prompt = f"""
You are a medical billing extraction assistant.

Extract structured data from the following definition.

Return STRICT JSON only in this format:

{{
  "serviceCodes": [],
  "diagnosisCodes": [],
  "revenueCodes": [],
  "modifier": "",
  "pos": [],
  "typeOfBill": "",
  "gender": "",
  "minAge": "",
  "maxAge": ""
}}

Rules:
- Expand service code ranges like "99242 to 99245" into list
- If nothing found return empty
- Return JSON only
- No explanation text

Definition:
{definition}
"""

    try:
        response = ollama.chat(
            model="llama3.1",
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0}
        )

        content = response["message"]["content"]

        # Remove ```json formatting if exists
        content = content.replace("```json", "").replace("```", "").strip()

        return json.loads(content)

    except Exception as e:
        print("⚠ LLM Error:", e)
        return {}


# -------------------------------
# Main Processing Function
# -------------------------------

def process_excel(input_file, output_file):

    df = pd.read_excel(input_file)
    output_rows = []

    for index, row in df.iterrows():

        service_category = row.get("SERVICE_CATEGORY_NAME", "")
        definition = row.get("DEFINITION", "")

        print(f"Processing: {service_category}")

        llm_data = call_llm(definition)

        if not llm_data:
            print("⚠ Skipping due to empty LLM response")
            continue

        # Extract fields safely
        service_codes = llm_data.get("serviceCodes", [])
        diagnosis_codes = llm_data.get("diagnosisCodes", [])
        revenue_codes = llm_data.get("revenueCodes", [])
        modifier = llm_data.get("modifier", "")
        pos_list = llm_data.get("pos", [])
        gender = llm_data.get("gender", "")
        min_age = llm_data.get("minAge", "")
        max_age = llm_data.get("maxAge", "")
        type_of_bill = llm_data.get("typeOfBill", "")

        # Expand service ranges manually (extra safety)
        service_codes = expand_service_ranges(service_codes)

        # Convert age range
        age_range = ""
        if min_age or max_age:
            age_range = f"{min_age}-{max_age}"

        # Safe string conversion
        diagnosis_string = safe_join(diagnosis_codes)
        revenue_string = safe_join(revenue_codes)
        pos_string = safe_join(pos_list)

        # Create multiple rows for each service code
        for code in map(str, service_codes):

            output_rows.append({
                "ServiceCategory": service_category,
                "ServiceCode": code,
                "RevenueCode": revenue_string,
                "Gender": str(gender),
                "Age": age_range,
                "DiagnosisCode": diagnosis_string,
                "POS": pos_string,
                "TypeOfBill": str(type_of_bill),
                "Modifier": str(modifier),
                "Minutes": 1,
                "Billed_Amnt": 100
            })

        # Small delay so Ollama doesn’t overload
        time.sleep(0.3)

    # Create output DataFrame
    output_df = pd.DataFrame(output_rows)

    # Write to Excel
    output_df.to_excel(output_file, index=False)

    print("\n✅ Output file created successfully!")


# -------------------------------
# Run Program
# -------------------------------

if __name__ == "__main__":
    process_excel("input.xlsx", "output.xlsx")
