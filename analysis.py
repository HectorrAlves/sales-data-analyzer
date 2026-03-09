import pandas as pd
import matplotlib.pyplot as plt

# ==========================
# Função para carregar dados
# ==========================
def load_data(file_name):
    print("Loading sales data...")
    data = pd.read_excel(file_name)
    return data


# ==========================
# Calcular métricas de vendas
# ==========================
def calculate_metrics(data):

    total_sales = data["valor"].sum()

    average_sales = data["valor"].mean()

    best_product = data.groupby("produto")["valor"].sum().idxmax()

    print("\n===== SALES METRICS =====")
    print("Total sales:", total_sales)
    print("Average sale:", average_sales)
    print("Best selling product:", best_product)


# ==========================
# Ranking de vendedores
# ==========================
def seller_ranking(data):

    ranking = data.groupby("vendedor")["valor"].sum().sort_values(ascending=False)

    print("\n===== SELLER RANKING =====")
    print(ranking)


# ==========================
# Gráfico de vendas
# ==========================
def sales_chart(data):

    sales_by_product = data.groupby("produto")["valor"].sum()

    sales_by_product.plot(kind="bar")

    plt.title("Sales by Product")
    plt.xlabel("Product")
    plt.ylabel("Total Sales")

    plt.tight_layout()

    plt.show()


# ==========================
# Função principal
# ==========================
def main():

    file_name = "vendas.xlsx"

    sales_data = load_data(file_name)

    calculate_metrics(sales_data)

    seller_ranking(sales_data)

    sales_chart(sales_data)


if __name__ == "__main__":
    main()
