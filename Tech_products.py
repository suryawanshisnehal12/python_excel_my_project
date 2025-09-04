import pandas as pd
import matplotlib.pyplot as plt


# Reading the Excel file (sheet "Tech") into a pandas DataFrame
df = pd.read_excel(r"D:\Data Anyaltics course\My Project\Tech_product_sales.xlsx", 
                   sheet_name="Tech", engine="openpyxl")

# Load dataset
print("File loaded:", df.shape)   # Prints the shape (rows, columns)

df.columns = df.columns.str.strip().str.lower()

# Handle missing values
df["comments"] = df["comments"].fillna("No comment") 
# Fill empty values in 'comments' column with "No comment"

# Convert datatypes
for c in ["product", "customer name", "comments"]:
    df[c] = df[c].astype(str).str.strip()

df["rating"] = df["rating"].astype(float)
# Convert rating to float for numeric operations
df["date"] = pd.to_datetime(df["date"], errors="coerce")

# Remove duplicates
df.drop_duplicates(inplace=True)


print("Average Rating Per Product:\n", df.groupby("product")["rating"].mean())
# Average rating per product
print("\nTop 6 reviews:\n", df.sort_values(by="rating", ascending=False).head(6))
# Top 10 reviews (sorted by rating, descending)

# Visualization
# Histogram-ratings
df["rating"].plot(kind="hist", bins=10, title="Histogram of Ratings")
plt.xlabel("Rating")
plt.show()

# Average rating per product as a bar chart
df.groupby("product")["rating"].mean().plot(kind="bar", title="Average Rating by Product")
plt.ylabel("Average Rating")
plt.show()

# Pie chart of average rating per product
df.groupby("product")["rating"].mean().plot(kind="pie",autopct='%1.1f%%', title="Rating by Product")
plt.ylabel(" ")  
plt.show()

#comparing product vs rating -scatter plot
df.reset_index().plot(kind="scatter", x="index", y="rating", title="Scatter: Product Index vs Rating")
plt.show()

# Cleaned excel file created 
df.to_excel("Cleaned_data.xlsx",index=False)




