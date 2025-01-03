from flask import Flask, request, render_template, send_file
import pandas as pd
import matplotlib.pyplot as plt
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'excel_file' not in request.files or 'quarter_files' not in request.files:
        return "No file part"

    excel_file = request.files['excel_file']
    quarter_files = request.files.getlist('quarter_files')

    if excel_file.filename == '' or not quarter_files:
        return "No selected file"

    # Save the uploaded files
    excel_file_path = os.path.join('uploads', excel_file.filename)
    excel_file.save(excel_file_path)

    quarter_file_paths = []
    for file in quarter_files:
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        quarter_file_paths.append(file_path)

    # Process the files
    executive_df = pd.read_excel(excel_file_path, usecols=["Department", "Executive Head"], engine='openpyxl')

    all_quarters_data = []
    quarter_files_dict = dict(zip(["JFM", "AMJ", "JAS", "OND"], quarter_file_paths))

    for quarter, file in quarter_files_dict.items():
        quarter_df = pd.read_csv(file, usecols=["Department", "Training Score", "Email"])
        merged_df = pd.merge(quarter_df, executive_df, on="Department", how="left")
        merged_df["Training Score"] = merged_df["Training Score"].apply(lambda x: "Yes" if pd.notna(x) else "No")
        merged_df["Quarter"] = quarter
        all_quarters_data.append(merged_df)

    yearly_data = pd.concat(all_quarters_data)

    executive_summary = []

    for executive, group in yearly_data.groupby("Executive Head"):
        total_members = group["Executive Head"].count()
        quarter_percentages = {}
        for quarter in ["JFM", "AMJ", "JAS", "OND"]:
            quarter_group = group[group["Quarter"] == quarter]
            total_in_quarter = quarter_group.shape[0]
            total_yes_in_quarter = quarter_group[quarter_group["Training Score"] == "Yes"].shape[0]
            percentage = (total_yes_in_quarter / total_in_quarter) * 100 if total_in_quarter > 0 else 0
            quarter_percentages[quarter] = percentage
        
        executive_summary.append({
            "ET Members": executive,
            "Total Members": total_members,
            **quarter_percentages
        })

    summary_df = pd.DataFrame(executive_summary)

    members_to_include = [
        "Ajoy Singh", "Ashwath Bhat", "Dylan Dias", "Manish Tiwari", 
        "Mrunali Majmudar Sathe", "Natwar Mall", "Rasesh Shah", 
        "Rohini Singh", "Sandeep Dutta", "Sankar SN Narayanan", 
        "Satish Raman", "Shailendra Singh"
    ]
    filtered_summary_df = summary_df[summary_df["ET Members"].isin(members_to_include)]

    grand_total_members = filtered_summary_df["Total Members"].sum()
    average_percentages = filtered_summary_df[["JFM", "AMJ", "JAS", "OND"]].mean()

    grand_total_row = pd.DataFrame({
        "ET Members": ["Grand Total"],
        "Total Members": [grand_total_members],
        "JFM": [average_percentages["JFM"]],
        "AMJ": [average_percentages["AMJ"]],
        "JAS": [average_percentages["JAS"]],
        "OND": [average_percentages["OND"]]
    })

    filtered_summary_df = pd.concat([filtered_summary_df, grand_total_row], ignore_index=True)

    output_file = "filtered_executive_summary.xlsx"
    filtered_summary_df.to_excel(output_file, index=False)

    filtered_summary_df.set_index("ET Members", inplace=True)
    filtered_summary_df.drop(columns=["Total Members"], inplace=True)
    filtered_summary_df.plot(kind="bar", figsize=(12, 8))
    plt.title("Training Score Percentage by Executive and Quarter")
    plt.ylabel("Percentage")
    plt.xlabel("Executive Head")
    plt.legend(title="Quarter")
    plt.yticks(range(0, 101, 10))
    plt.tight_layout()

    chart_output_file = "filtered_executive_summary_chart.png"
    plt.savefig(chart_output_file)
    plt.close()

    return render_template('result.html', excel_file=output_file, chart_file=chart_output_file)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)