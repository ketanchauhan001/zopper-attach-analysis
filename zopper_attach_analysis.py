import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


DATA_FILE = "Jumbo & Company_ Attach % .xls"
OUTPUT_FILE = "zopper_attach_analysis_output.xlsx"
MONTHS = ["Aug", "Sep", "Oct", "Nov", "Dec"]


# DATA LOADING & CLEANING
def load_data(file_path: str) -> pd.DataFrame:
    """Read Excel file and convert month columns into long format."""

    df = pd.read_excel(file_path, engine="xlrd")
    print("Data loaded successfully")
    print("Columns available:", list(df.columns))

    # Convert wide monthly data into long format
    df_long = df.melt(
        id_vars=["Branch", "Store_Name"],
        value_vars=MONTHS,
        var_name="Month",
        value_name="Attach_Percent",
    )

    # Convert fraction to percentage
    df_long["Attach_Percent"] = df_long["Attach_Percent"] * 100

    # Keep months in correct order
    df_long["Month"] = pd.Categorical(
        df_long["Month"], categories=MONTHS, ordered=True
    )

    return df_long



# ANALYSIS FUNCTIONS
def analyze_by_month(df: pd.DataFrame) -> pd.DataFrame:
    """Average Attach % month-wise."""
    return df.groupby("Month", as_index=False)["Attach_Percent"].mean()


def analyze_by_branch(df: pd.DataFrame) -> pd.DataFrame:
    """Average Attach % branch-wise."""
    return (
        df.groupby("Branch", as_index=False)["Attach_Percent"]
        .mean()
        .sort_values("Attach_Percent", ascending=False)
    )


def analyze_store_performance(df: pd.DataFrame) -> pd.DataFrame:
    """Average Attach % for each store."""
    return (
        df.groupby(["Branch", "Store_Name"], as_index=False)["Attach_Percent"]
        .mean()
    )




# STORE CATEGORY LOGIC
def store_category(value: float) -> str:
    if value >= 30:
        return "High Performer"
    elif value >= 15:
        return "Medium Performer"
    return "Low Performer"




# JANUARY PREDICTION (SIMPLE AVERAGE METHOD)
def predict_january_attach(df: pd.DataFrame) -> pd.DataFrame:
    """Predict January Attach % using last 3 months average."""

    prediction = (
        df.sort_values("Month")
        .groupby(["Branch", "Store_Name"])
        .tail(3)
        .groupby(["Branch", "Store_Name"], as_index=False)["Attach_Percent"]
        .mean()
        .rename(columns={"Attach_Percent": "Predicted_Jan_Attach%"})
    )

    return prediction




# VISUALIZATION
def plot_month_trend(df: pd.DataFrame):
    plt.figure()
    sns.lineplot(data=df, x="Month", y="Attach_Percent", marker="o")
    plt.title("Average Attach % by Month")
    plt.tight_layout()
    plt.show()


def plot_branch_performance(df: pd.DataFrame):
    plt.figure(figsize=(10, 5))
    sns.barplot(data=df, x="Branch", y="Attach_Percent")
    plt.title("Average Attach % by Branch")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


def plot_store_distribution(df: pd.DataFrame):
    plt.figure()
    sns.scatterplot(
        data=df, x="Store_Name", y="Attach_Percent", hue="Category"
    )
    plt.title("Store-wise Attach % Distribution")
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.show()




# MAIN WORKFLOW
def main():
    print("Starting Attach % analysis...")

    data = load_data(DATA_FILE)

    month_analysis = analyze_by_month(data)
    branch_analysis = analyze_by_branch(data)
    store_analysis = analyze_store_performance(data)

    store_analysis["Category"] = store_analysis["Attach_Percent"].apply(store_category)

    january_prediction = predict_january_attach(data)

    print("Creating charts...")
    plot_month_trend(month_analysis)
    plot_branch_performance(branch_analysis)
    plot_store_distribution(store_analysis)

    print("Saving results to Excel...")
    with pd.ExcelWriter(OUTPUT_FILE) as writer:
        data.to_excel(writer, "Cleaned_Data", index=False)
        month_analysis.to_excel(writer, "Month_Wise", index=False)
        branch_analysis.to_excel(writer, "Branch_Wise", index=False)
        store_analysis.to_excel(writer, "Store_Performance", index=False)
        january_prediction.to_excel(writer, "January_Prediction", index=False)

    print("Analysis completed successfully")


if __name__ == "__main__":
    main()
