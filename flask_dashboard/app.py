from flask import Flask, render_template, request, redirect, url_for, send_file, Response, jsonify
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from prophet import Prophet
from tqdm import tqdm
import os
import io
import base64
import time
from openpyxl import Workbook

def load_data(file_path):
    return pd.read_excel(file_path)

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def plot_to_img(fig):
    img = io.BytesIO()
    fig.savefig(img, format='png', bbox_inches='tight')  # Ensure plots are not cropped
    img.seek(0)
    return base64.b64encode(img.getvalue()).decode()

def generate_forecast(df, progress_callback=None):
    df["Date"] = pd.to_datetime(df["Date"])
    
    # Check if Sales Name column exists
    if "Sales Name" in df.columns:
        # Use Sales-Customer-Item triplets
        grouping_cols = ["Sales Name", "Customer Name", "Item Name"]
        triplets = df.groupby(grouping_cols).size().reset_index().drop(0, axis=1)
        print(f"Using Sales-Customer-Item triplets: {len(triplets)} combinations")
    else:
        # Fallback to Customer-Item pairs
        grouping_cols = ["Customer Name", "Item Name"]
        triplets = df.groupby(grouping_cols).size().reset_index().drop(0, axis=1)
        print(f"Using Customer-Item pairs: {len(triplets)} combinations")
    
    all_forecasts = []
    total_triplets = len(triplets)

    for idx, row in enumerate(tqdm(triplets.iterrows(), total=total_triplets, desc="Training Prophet Model")):
        row_data = row[1]
        
        # Extract grouping values
        if "Sales Name" in df.columns:
            sales_name = row_data["Sales Name"]
            customer = row_data["Customer Name"]
            item = row_data["Item Name"]
            # Filter for this specific triplet
            df_filtered = df[
                (df["Sales Name"] == sales_name) &
                (df["Customer Name"] == customer) & 
                (df["Item Name"] == item)
            ][["Date", "Quantity"]]
        else:
            customer = row_data["Customer Name"]
            item = row_data["Item Name"]
            # Filter for this specific pair
            df_filtered = df[
                (df["Customer Name"] == customer) & 
                (df["Item Name"] == item)
            ][["Date", "Quantity"]]
        
        # Resample to monthly
        df_filtered = df_filtered.resample("ME", on="Date").sum().reset_index()
        
        if len(df_filtered) < 10:
            continue

        # Add required columns to df_filtered
        if "Sales Name" in df.columns:
            df_filtered["Sales Name"] = sales_name
        df_filtered["Customer Name"] = customer
        df_filtered["Item Name"] = item
        df_filtered["Type"] = "Actual"
        df_filtered = df_filtered.rename(columns={"Quantity": "Actual Quantity"})

        # Prepare data for Prophet
        prophet_df = df_filtered.rename(columns={"Date": "ds", "Actual Quantity": "y"})
        model = Prophet()
        model.fit(prophet_df)

        # Generate forecast
        future = model.make_future_dataframe(periods=12, freq='ME')
        forecast = model.predict(future)
        forecast = forecast[forecast["yhat"] >= 0]
        
        # Add metadata to forecast
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

        # Merge actual and forecast data
        merged_df = pd.concat([df_filtered, forecast], ignore_index=True)
        all_forecasts.append(merged_df)

        # Emit progress update
        if progress_callback:
            progress = int((idx + 1) / total_triplets * 100)
            progress_callback(progress)

    return pd.concat(all_forecasts) if all_forecasts else pd.DataFrame()

def plot_top_customers(df):
    exchange_rate = 16000  # 1 USD = 16,000 IDR
    df["Amount_in_IDR"] = df["Amount"].apply(lambda x: x * exchange_rate if x < 10000 else x)

    # Group by Customer Name and aggregate total Quantity & Amount in Rupiah
    top_customers = df.groupby("Customer Name").agg({"Quantity": "sum", "Amount_in_IDR": "sum"}).reset_index()

    # Get top 5 customers by Quantity
    top_5_qty = top_customers.sort_values("Quantity", ascending=False).head(5)

    # Get top 5 customers by Rupiah Value
    top_5_value = top_customers.sort_values("Amount_in_IDR", ascending=False).head(5)

    # Create subplots
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))

    # Plot Top 5 Customers by Quantity
    sns.barplot(x="Quantity", y="Customer Name", data=top_5_qty, ax=axes[0], palette="Blues_r")
    axes[0].set_title("Top 5 Customers by Quantity")
    axes[0].set_xlabel("Total Quantity")
    axes[0].set_ylabel("Customer Name")

    # Add numbers on bars
    for index, value in enumerate(top_5_qty["Quantity"]):
        axes[0].text(value, index, f"{value:,.0f}", va='center')

    # Plot Top 5 Customers by Sales Value (Converted to IDR)
    sns.barplot(x="Amount_in_IDR", y="Customer Name", data=top_5_value, ax=axes[1], palette="Greens_r")
    axes[1].set_title("Top 5 Customers by Sales Value (in IDR)")
    axes[1].set_xlabel("Total Amount (IDR)")
    axes[1].set_ylabel("")

    # Add numbers on bars with currency format
    for index, value in enumerate(top_5_value["Amount_in_IDR"]):
        axes[1].text(value, index, f"Rp {value:,.0f}", va='center')

    # Adjust layout
    plt.tight_layout()
    return plot_to_img(fig)

def plot_top_cities(df):
    # Count unique 'Document Number' per city
    city_counts = df.groupby("City")["Document Number"].nunique().reset_index()
    city_counts.columns = ["City", "Count"]  # Rename columns

    # Get top 10 cities
    top_cities = city_counts.nlargest(10, "Count")

    # Plot
    plt.figure(figsize=(10, 6))
    fig = plt.gcf()  # Get the current figure
    ax = sns.barplot(x="Count", y="City", data=top_cities, palette="Blues_r")

    # Add numbers on bars
    for index, value in enumerate(top_cities["Count"]):
        ax.text(value, index, f"{value:,.0f}", va='center')

    plt.title("Top 10 Cities by Unique Shipment Count")
    plt.xlabel("Number of Unique Shipments")
    plt.ylabel("City")
    return plot_to_img(fig)

def plot_top_items(df):
    # Group by Item Name and sum Quantity
    item_sales = df.groupby("Item Name")["Quantity"].sum().reset_index()

    # Get top 10 items
    top_items = item_sales.sort_values("Quantity", ascending=False).head(10)

    # Plot
    plt.figure(figsize=(10, 6))
    fig = plt.gcf()  # Get the current figure
    ax = sns.barplot(x="Quantity", y="Item Name", data=top_items, palette="Greens_r")
    
    # Add numbers on bars
    for index, value in enumerate(top_items["Quantity"]):
        ax.text(value, index, f"{value:,.0f}", va='center')

    plt.title("Top 10 Most Purchased Items")
    plt.xlabel("Total Quantity Purchased")
    plt.ylabel("Item Name")
    return plot_to_img(fig)

def plot_top_salespeople(df):
    # Define exchange rates (converting everything to IDR)
    exchange_rates = {
        "Rupiah": 1,       # IDR remains unchanged
        "US Dollar": 16000    # 1 USD = 16,000 IDR
    }

    # Convert all amounts to IDR
    df["Amount_Converted"] = df.apply(lambda row: row["Amount"] * exchange_rates.get(row["Currency"], 1), axis=1)

    # Group by Sales Name and sum Quantity & Amount (in IDR)
    sales_performance = df.groupby("Sales Name").agg({"Quantity": "sum", "Amount_Converted": "sum"}).reset_index()

    # Get top 10 salespeople
    top_sales = sales_performance.sort_values("Amount_Converted", ascending=False).head(10)

    # Plot with correct amounts
    plt.figure(figsize=(10, 6))
    fig = plt.gcf()  # Get the current figure
    ax = sns.barplot(x="Amount_Converted", y="Sales Name", data=top_sales, palette="Reds_r")

    # Add numbers on bars (formatted in Rupiah)
    for index, value in enumerate(top_sales["Amount_Converted"]):
        ax.text(value, index, f"Rp {value:,.0f}", va='center')

    plt.title("Top 10 Salespeople by Total Sales (Converted to IDR)")
    plt.xlabel("Total Sales Amount (in IDR)")
    plt.ylabel("Salesperson")
    return plot_to_img(fig)

def get_quarterly_customer_activity(df):
    # Convert Date column to datetime format if not already
    df["Date"] = pd.to_datetime(df["Date"])

    # Create a column for quarterly periods (Year + Quarter)
    df["Quarter"] = df["Date"].dt.to_period("Q")

    # Get all unique quarters in the dataset
    all_quarters = df["Quarter"].unique()

    # Convert PeriodArray to a sorted list of strings
    all_quarters = sorted([str(q) for q in all_quarters])

    # Get all unique customers in the dataset
    all_customers = set(df["Customer Name"].unique())

    # Initialize a dictionary to track active and inactive customers per quarter
    quarterly_activity = []

    # Track cumulative active customers up to each quarter
    cumulative_active_customers = set()

    for quarter in all_quarters:
        # Convert the quarter back to a Period object for filtering
        quarter_period = pd.Period(quarter)

        # Get customers active in the current quarter
        current_active = set(df[df["Quarter"] == quarter_period]["Customer Name"].unique())

        # Update cumulative active customers
        cumulative_active_customers.update(current_active)

        # Get inactive customers (customers who have transacted before but not in this quarter)
        inactive_customers = cumulative_active_customers - current_active

        # Append the results for the current quarter
        quarterly_activity.append({
            "quarter": str(quarter),
            "active_customers": list(current_active),  # Include all active customers
            "inactive_customers": list(inactive_customers)  # Include all inactive customers
        })

    return quarterly_activity

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            return redirect(url_for("confirm", filename=file.filename))  # Redirect to confirm page
    return render_template("index.html")

@app.route("/confirm/<filename>")
def confirm(filename):
    return render_template("confirm.html", filename=filename)

@app.route("/loading/<filename>")
def loading(filename):
    return render_template("loading.html", filename=filename)

@app.route("/progress/<filename>")
def progress(filename):
    def generate_progress():
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        df = load_data(file_path)
        
        # Check if Sales Name column exists for proper counting
        if "Sales Name" in df.columns:
            total_triplets = len(df.groupby(["Sales Name", "Customer Name", "Item Name"]).size())
        else:
            total_triplets = len(df.groupby(["Customer Name", "Item Name"]).size())

        def progress_callback(progress):
            yield f"data: {progress}\n\n"

        forecast_result = generate_forecast(df, progress_callback=progress_callback)

        # Save forecast to CSV
        forecast_csv_path = os.path.join(UPLOAD_FOLDER, "forecast_result.csv")
        forecast_result.to_csv(forecast_csv_path, index=False)

        yield "data: 100\n\n"  # Training complete

    return Response(generate_progress(), mimetype="text/event-stream")

@app.route("/dashboard/<filename>")
def dashboard(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    df = load_data(file_path)

    # Check if training was skipped
    train_model = request.args.get("train", "true").lower() == "true"

    # Generate plots (only if training was not skipped)
    if train_model:
        top_customers_img = plot_top_customers(df)
        top_cities_img = plot_top_cities(df)
        top_items_img = plot_top_items(df)
        top_sales_img = plot_top_salespeople(df)
    else:
        # Generate plots without training
        top_customers_img = plot_top_customers(df)
        top_cities_img = plot_top_cities(df)
        top_items_img = plot_top_items(df)
        top_sales_img = plot_top_salespeople(df)

    # Get quarterly customer activity
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
    forecast_csv_path = os.path.join(UPLOAD_FOLDER, "forecast_result.csv")
    return send_file(forecast_csv_path, as_attachment=True, download_name="forecast_result.csv")

@app.route("/download_quarterly_activity/<filename>")
def download_quarterly_activity(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    df = load_data(file_path)

    # Get quarterly customer activity data
    quarterly_activity = get_quarterly_customer_activity(df)

    # Create a new Excel workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Quarterly Customer Activity"

    # Add headers to the worksheet
    ws.append(["Quarter", "Active Customers", "Inactive Customers"])

    # Add data to the worksheet
    for quarter in quarterly_activity:
        active_customers = ", ".join(quarter["active_customers"])
        inactive_customers = ", ".join(quarter["inactive_customers"])
        ws.append([quarter["quarter"], active_customers, inactive_customers])

    # Save the workbook to a file
    excel_path = os.path.join(UPLOAD_FOLDER, "quarterly_activity.xlsx")
    wb.save(excel_path)

    # Return the Excel file as a download
    return send_file(excel_path, as_attachment=True, download_name="quarterly_activity.xlsx")

if __name__ == "__main__":
    import webbrowser
    from threading import Timer

    # Open the browser after a delay
    def open_browser():
        webbrowser.open_new("http://127.0.0.1:5000")

    # Wait 3 seconds before opening the browser
    Timer(3, open_browser).start()

    # Run the Flask app
    app.run()
