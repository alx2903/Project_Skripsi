from flask import Flask, render_template, request, redirect, url_for, send_file, Response, jsonify
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from prophet import Prophet
from tqdm import tqdm
import os
import io
import base64
import time
from openpyxl import Workbook
import threading
import json

def load_data(file_path):
    return pd.read_excel(file_path)

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Global dictionary to track training status
training_status = {}

def plot_to_img(fig):
    img = io.BytesIO()
    fig.savefig(img, format='png', bbox_inches='tight')
    img.seek(0)
    plt.close(fig)
    return base64.b64encode(img.getvalue()).decode()

def generate_forecast(df, filename, progress_callback=None):
    df["Date"] = pd.to_datetime(df["Date"])

    if "Sales Name" in df.columns:
        grouping_cols = ["Sales Name", "Customer Name", "Item Name"]
        triplets = df.groupby(grouping_cols).size().reset_index().drop(0, axis=1)
        print(f"Using Sales-Customer-Item triplets: {len(triplets)} combinations")
    else:
        grouping_cols = ["Customer Name", "Item Name"]
        triplets = df.groupby(grouping_cols).size().reset_index().drop(0, axis=1)
        print(f"Using Customer-Item pairs: {len(triplets)} combinations")

    all_forecasts = []
    total_triplets = len(triplets)

    for idx, row in enumerate(triplets.iterrows()):
        row_data = row[1]

        if "Sales Name" in df.columns:
            sales_name = row_data["Sales Name"]
            customer = row_data["Customer Name"]
            item = row_data["Item Name"]
            df_filtered = df[
                (df["Sales Name"] == sales_name) &
                (df["Customer Name"] == customer) &
                (df["Item Name"] == item)
            ][["Date", "Quantity"]]
        else:
            customer = row_data["Customer Name"]
            item = row_data["Item Name"]
            df_filtered = df[
                (df["Customer Name"] == customer) &
                (df["Item Name"] == item)
            ][["Date", "Quantity"]]

        df_filtered = df_filtered.resample("M", on="Date").sum().reset_index()

        if len(df_filtered) < 10:
            continue

        if "Sales Name" in df.columns:
            df_filtered["Sales Name"] = sales_name
        df_filtered["Customer Name"] = customer
        df_filtered["Item Name"] = item
        df_filtered["Type"] = "Actual"
        df_filtered = df_filtered.rename(columns={"Quantity": "Actual Quantity"})

        prophet_df = df_filtered.rename(columns={"Date": "ds", "Actual Quantity": "y"})
        model = Prophet()
        model.fit(prophet_df)

        future = model.make_future_dataframe(periods=12, freq='M')
        forecast = model.predict(future)
        forecast = forecast[forecast["yhat"] >= 0]

        if "Sales Name" in df.columns:
            forecast["Sales Name"] = sales_name
        forecast["Customer Name"] = customer
        forecast["Item Name"] = item
        forecast["Type"] = "Forecast"
        forecast = forecast.rename(columns={
            "ds": "Date",
            "yhat": "Predicted Quantity",
            "yhat_lower": "yhat_lower",
            "yhat_upper": "yhat_upper"
        })

        merged_df = pd.concat([df_filtered, forecast], ignore_index=True)
        all_forecasts.append(merged_df)

        # Update progress
        progress = int((idx + 1) / total_triplets * 100)
        training_status[filename]["progress"] = progress
        if progress_callback:
            progress_callback(progress)

    result = pd.concat(all_forecasts) if all_forecasts else pd.DataFrame()

    # Save forecast
    forecast_csv_path = os.path.join(UPLOAD_FOLDER, f"forecast_{filename}.csv")
    result.to_csv(forecast_csv_path, index=False)

    # Mark as complete
    training_status[filename]["complete"] = True
    training_status[filename]["progress"] = 100

    return result

def plot_top_customers(df):
    exchange_rate = 16000
    df["Amount_in_IDR"] = df["Amount"].apply(lambda x: x * exchange_rate if x < 10000 else x)

    top_customers = df.groupby("Customer Name").agg({"Quantity": "sum", "Amount_in_IDR": "sum"}).reset_index()
    top_5_qty = top_customers.sort_values("Quantity", ascending=False).head(5)
    top_5_value = top_customers.sort_values("Amount_in_IDR", ascending=False).head(5)

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))

    sns.barplot(x="Quantity", y="Customer Name", data=top_5_qty, ax=axes[0], palette="Blues_r")
    axes[0].set_title("Top 5 Customers by Quantity")
    axes[0].set_xlabel("Total Quantity")
    axes[0].set_ylabel("Customer Name")

    for index, value in enumerate(top_5_qty["Quantity"]):
        axes[0].text(value, index, f"{value:,.0f}", va='center')

    sns.barplot(x="Amount_in_IDR", y="Customer Name", data=top_5_value, ax=axes[1], palette="Greens_r")
    axes[1].set_title("Top 5 Customers by Sales Value (in IDR)")
    axes[1].set_xlabel("Total Amount (IDR)")
    axes[1].set_ylabel("")

    for index, value in enumerate(top_5_value["Amount_in_IDR"]):
        axes[1].text(value, index, f"Rp {value:,.0f}", va='center')

    plt.tight_layout()
    return plot_to_img(fig)

def plot_top_cities(df):
    city_counts = df.groupby("City")["Document Number"].nunique().reset_index()
    city_counts.columns = ["City", "Count"]
    top_cities = city_counts.nlargest(10, "Count")

    plt.figure(figsize=(10, 6))
    fig = plt.gcf()
    ax = sns.barplot(x="Count", y="City", data=top_cities, palette="Blues_r")

    for index, value in enumerate(top_cities["Count"]):
        ax.text(value, index, f"{value:,.0f}", va='center')

    plt.title("Top 10 Cities by Unique Shipment Count")
    plt.xlabel("Number of Unique Shipments")
    plt.ylabel("City")
    return plot_to_img(fig)

def plot_top_items(df):
    item_sales = df.groupby("Item Name")["Quantity"].sum().reset_index()
    top_items = item_sales.sort_values("Quantity", ascending=False).head(10)

    plt.figure(figsize=(10, 6))
    fig = plt.gcf()
    ax = sns.barplot(x="Quantity", y="Item Name", data=top_items, palette="Greens_r")

    for index, value in enumerate(top_items["Quantity"]):
        ax.text(value, index, f"{value:,.0f}", va='center')

    plt.title("Top 10 Most Purchased Items")
    plt.xlabel("Total Quantity Purchased")
    plt.ylabel("Item Name")
    return plot_to_img(fig)

def plot_top_salespeople(df):
    exchange_rates = {
        "Rupiah": 1,
        "US Dollar": 16000
    }

    df["Amount_Converted"] = df.apply(lambda row: row["Amount"] * exchange_rates.get(row["Currency"], 1), axis=1)
    sales_performance = df.groupby("Sales Name").agg({"Quantity": "sum", "Amount_Converted": "sum"}).reset_index()
    top_sales = sales_performance.sort_values("Amount_Converted", ascending=False).head(10)

    plt.figure(figsize=(10, 6))
    fig = plt.gcf()
    ax = sns.barplot(x="Amount_Converted", y="Sales Name", data=top_sales, palette="Reds_r")

    for index, value in enumerate(top_sales["Amount_Converted"]):
        ax.text(value, index, f"Rp {value:,.0f}", va='center')

    plt.title("Top 10 Salespeople by Total Sales (Converted to IDR)")
    plt.xlabel("Total Sales Amount (in IDR)")
    plt.ylabel("Salesperson")
    return plot_to_img(fig)

def get_quarterly_customer_activity(df):
    df["Date"] = pd.to_datetime(df["Date"])
    df["Quarter"] = df["Date"].dt.to_period("Q")
    all_quarters = sorted([str(q) for q in df["Quarter"].unique()])
    quarterly_activity = []
    cumulative_active_customers = set()

    for quarter in all_quarters:
        quarter_period = pd.Period(quarter)
        current_active = set(df[df["Quarter"] == quarter_period]["Customer Name"].unique())
        cumulative_active_customers.update(current_active)
        inactive_customers = cumulative_active_customers - current_active

        quarterly_activity.append({
            "quarter": str(quarter),
            "active_customers": list(current_active),
            "inactive_customers": list(inactive_customers)
        })

    return quarterly_activity

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            return redirect(url_for("confirm", filename=file.filename))
    return render_template("index.html")

@app.route("/confirm/<filename>")
def confirm(filename):
    return render_template("confirm.html", filename=filename)

@app.route("/loading/<filename>")
def loading(filename):
    return render_template("loading.html", filename=filename)

@app.route("/start_training/<filename>")
def start_training(filename):
    """Start training in background thread"""
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    # Initialize training status
    training_status[filename] = {
        "complete": False,
        "progress": 0,
        "error": None
    }

    # Start training in background thread
    def train_in_background():
        try:
            df = load_data(file_path)
            generate_forecast(df, filename)
        except Exception as e:
            training_status[filename]["error"] = str(e)
            training_status[filename]["complete"] = True

    thread = threading.Thread(target=train_in_background)
    thread.daemon = True
    thread.start()

    return jsonify({"status": "started"})

@app.route("/check_training_status/<filename>")
def check_training_status(filename):
    """Check training progress"""
    status = training_status.get(filename, {"complete": False, "progress": 0, "error": None})
    return jsonify(status)

@app.route("/dashboard/<filename>")
def dashboard(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    df = load_data(file_path)

    train_model = request.args.get("train", "true").lower() == "true"

    top_customers_img = plot_top_customers(df)
    top_cities_img = plot_top_cities(df)
    top_items_img = plot_top_items(df)
    top_sales_img = plot_top_salespeople(df)

    quarterly_activity = get_quarterly_customer_activity(df)

    return render_template("dashboard.html", filename=filename,
                           top_customers_img=top_customers_img,
                           top_cities_img=top_cities_img,
                           top_items_img=top_items_img,
                           top_sales_img=top_sales_img,
                           quarterly_activity=quarterly_activity,
                           train_model=train_model)

@app.route("/loading_no_training/<filename>")
def loading_no_training(filename):
    return render_template("loading_no_training.html", filename=filename)

@app.route("/download_forecast/<filename>")
def download_forecast(filename):
    forecast_csv_path = os.path.join(UPLOAD_FOLDER, f"forecast_{filename}.csv")
    if not os.path.exists(forecast_csv_path):
        forecast_csv_path = os.path.join(UPLOAD_FOLDER, "forecast_result.csv")
    return send_file(forecast_csv_path, as_attachment=True, download_name="forecast_result.csv")

@app.route("/download_quarterly_activity/<filename>")
def download_quarterly_activity(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    df = load_data(file_path)

    quarterly_activity = get_quarterly_customer_activity(df)

    wb = Workbook()
    ws = wb.active
    ws.title = "Quarterly Customer Activity"

    ws.append(["Quarter", "Active Customers", "Inactive Customers"])

    for quarter in quarterly_activity:
        active_customers = ", ".join(quarter["active_customers"])
        inactive_customers = ", ".join(quarter["inactive_customers"])
        ws.append([quarter["quarter"], active_customers, inactive_customers])

    excel_path = os.path.join(UPLOAD_FOLDER, "quarterly_activity.xlsx")
    wb.save(excel_path)

    return send_file(excel_path, as_attachment=True, download_name="quarterly_activity.xlsx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
