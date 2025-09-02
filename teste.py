import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

def gera_gráfico_pila():
    # Dados
    categorias = ["E", "P", "C"]
    desejado = [4, 4, 5]
    apurado = [3, 3, 4]
    gap = [a - b for a, b in zip( apurado, desejado)]

    x = np.arange(len(categorias))
    largura = 0.25

    fig, ax = plt.subplots(figsize=(8, 4))

    # Barras
    b1 = ax.bar(x - largura, desejado, largura, label="Desejado", color="#4472C4")
    b2 = ax.bar(x, apurado, largura, label="Apurado", color="#ED7D31")
    b3 = ax.bar(x + largura, gap, largura, label="GAP", color="gray")

    # Título
    ax.set_title("Comparativo Pilares", fontsize=14, fontweight="bold")

    # Eixo X
    ax.set_xticks(x)
    ax.set_xticklabels(categorias)

    # Eixo Y com margem para não cortar rótulos
    y_min = min(min(desejado), min(apurado), min(gap)) - 1.5
    y_max = max(max(desejado), max(apurado), max(gap)) + 1.5
    ax.set_ylim(y_min, y_max)

    # Linha zero
    ax.axhline(0, color="black", linewidth=0.8)

    # Legenda fora do gráfico (lado direito)
    ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))

    # Função para rótulos
    def autolabel(barras):
        for barra in barras:
            altura = barra.get_height()
            ax.annotate(f'{altura:.0f}',
                        xy=(barra.get_x() + barra.get_width() / 2, altura),
                        xytext=(0, 5 if altura >= 0 else -10),
                        textcoords="offset points",
                        ha='center', va='bottom')

    autolabel(b1)
    autolabel(b2)
    autolabel(b3)

    plt.tight_layout()
    plt.savefig("grafico.png", dpi=300, bbox_inches="tight")
