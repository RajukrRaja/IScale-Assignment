import pandas as pd
import numpy as np
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

def main():
    """
    The main function to load data, explore them, and present results.
    """
    # Hide the root window
    Tk().withdraw()

    # Ask for the expected incoming Excel file
    file_path = askopenfilename(title="Select Excel file")

    if not file_path:
        print("No file selected. Exiting.")
        return

    # Load data from the given path
    try:
        data = pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"File not found at {file_path}. Please check the path and try again.")
        return
    except Exception as e:
        print(f"Error while reading the file: {e}")
        return

    # Convert to datetime data types for time columns
    data['slot_start_time'] = pd.to_datetime(data['slot_start_time'], errors='coerce')
    data['payment_time'] = pd.to_datetime(data['payment_time'], errors='coerce')

    # 1. Calculate 3-Day Conversion Rates
    within_3_days = data[(data['payment_time'] - data['slot_start_time']).dt.days <= 3]
    conversion_3_days = within_3_days.groupby(['funnel', 'medicalconditionflag']).size() / data.groupby(['funnel', 'medicalconditionflag']).size()
    conversion_3_days = conversion_3_days.reset_index(name='conversion_rate_3_days')

    # 2. Calculate 7-Day Conversion Rates
    within_7_days = data[(data['payment_time'] - data['slot_start_time']).dt.days <= 7]
    conversion_7_days = within_7_days.groupby(['funnel', 'medicalconditionflag']).size() / data.groupby(['funnel', 'medicalconditionflag']).size()
    conversion_7_days = conversion_7_days.reset_index(name='conversion_rate_7_days')

    # 3. Analyze Hourly Distributions
    slot_hour_distribution = data['slot_start_time'].dt.hour.value_counts().sort_index()
    payment_hour_distribution = data['payment_time'].dt.hour.value_counts().sort_index()

    # 4. Coach & Funnel Performance Evaluation
    coach_performance = data.groupby('target_class')['payment_time'].count().reset_index(name='conversions')
    funnel_performance = data.groupby('funnel')['payment_time'].count().reset_index(name='conversions')

    # 5. Review Medical Condition Impact
    medical_condition_impact = data.groupby(['medicalconditionflag', 'funnel'])['payment_time'].count().reset_index(name='conversions')

    # Visualization
    # 1. Conversion Rates Visualization using Seaborn
    sns.barplot(x='funnel', y='conversion_rate_3_days', hue='medicalconditionflag', data=conversion_3_days)
    plt.title('3-Day Conversion Rates by Funnel and Medical Condition')
    plt.show()

    sns.barplot(x='funnel', y='conversion_rate_7_days', hue='medicalconditionflag', data=conversion_7_days)
    plt.title('7-Day Conversion Rates by Funnel and Medical Condition')
    plt.show()

    # 2. Hourly Distribution Visualization using Matplotlib
    plt.figure(figsize=(10, 5))
    plt.subplot(1, 2, 1)
    slot_hour_distribution.plot(kind='bar', color='skyblue')
    plt.title('Slot Hour Distribution')

    plt.subplot(1, 2, 2)
    payment_hour_distribution.plot(kind='bar', color='lightgreen')
    plt.title('Payment Hour Distribution')
    plt.tight_layout()
    plt.show()

    # 3. Funnel Performance using Plotly
    fig = px.bar(funnel_performance, x='funnel', y='conversions', title="Funnel Performance")
    fig.show()

    # 4. Coach Performance using Plotly
    fig = px.pie(coach_performance, values='conversions', names='target_class', title="Coach Performance")
    fig.show()

if __name__ == "__main__":
    main()
