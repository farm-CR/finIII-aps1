import pandas as pd
import numpy as np
from datetime import datetime as dt
import plotly.express as px
import os

class Strategy:
    def __init__(self, strategy):

        self.strategy = strategy.split(".xlsx")[0]

        self.df = (pd.read_excel(f"input/{self.strategy}.xlsx")
        .assign(
            #Cálculo do valor presente do Ativo Risk Free
            Vcto = lambda _: (_.Vcto - dt.now()).apply(lambda __: __.days), 
            Strike = lambda _: np.where(_.Tipo == "Ativo Rf", _.Valor * (1 + _.Strike) ** (_.Vcto / 360), _.Strike),

            #Cálculo do custo de cada ativo (sinal contrário do fluxo de caixa no t0)
            Custo = lambda _: _.Valor * _.Qtde
        ).drop("Vcto", axis = 1)

        #Criação de valores de St para cada ativo
        .assign(merge = 0)
        .pipe(lambda _: 
            _.merge(pd.DataFrame({
                "St": np.linspace(_.Strike.values[0] / 1.1, _.Strike.values[-1] * 1.1, 10000), 
                "merge": np.zeros(10000)}),
            on = "merge", how = "outer")
        )
        .drop("merge", axis = 1)
        
        #Cálculo do payoff para cada valor de St
        .assign(
            Payoff = lambda _: np.where(
                _.Tipo == "Call", #payoff da call
                np.maximum(_.St - _.Strike, 0) * _.Qtde,
                np.where(
                    _.Tipo == "Put", #payoff da put
                    np.maximum(_.Strike - _.St, 0) * _.Qtde,
                    np.where(
                        _.Tipo == "Ação", 
                        _.St * _.Qtde, #payoff da ação
                        _.Strike * _.Qtde #payoff do Risk Free
                    )
                )
            ),
            #Cálculo do resultado para cada valor de St
            Resultado = lambda _: _.Payoff - _.Custo
        )

        #Formatação do dataframe fazer o gráfico
        .assign(Ativo = lambda _: "[" + _.Qtde.astype(str) + "] " + _.Tipo.astype(str) + " (" + _.Ativo.astype(str)  + ")")
        .filter(["Ativo", "St", "Payoff", "Resultado"])
        .pipe(lambda _: 
            pd.concat([
                (_ #Dataframe com o payoff de cada ativo
                .groupby(["St"]).sum()
                .reset_index()
                .melt("St")
                .assign(variable = lambda _: _.variable.astype(str) + "_Conjunto",
                        dashed = False)  
                ), (_ #Dataframe com o payoff e resultado conjuntos
                .drop("Resultado", axis = 1)
                .assign(dashed = True)
                .rename({"Ativo": "variable", "Payoff": "value"}, axis = 1)
                )
            ]))
        )
    
    #Exporta o gráfico para a pasta "output"
    def plot(self):

        px.line(
            self.df, x = "St", y = "value", color = "variable", line_dash = "dashed", 
            labels = {
                "St": "Valor do Ativo-Objeto na data de vencimento (St)",
                "value": "Payoff / Resultado",
                "variable": "Ativo",
                "dashed": "linha pontilhada"
            }, title = self.strategy.replace("-", " ").capitalize()
        ).write_html(f"output/{self.strategy}.html")


    #Exporta o relatório para a pasta "output"   
    def report(self):

        tabela = (self.df
            .query("variable in ['Payoff_Conjunto', 'Resultado_Conjunto']")
            .pivot_table(index = "St", columns= ["variable"], values = "value")
            .iloc[[index for index in range(10000) if index % (10000 / 10) == 0]+[-1]]
            .rename_axis("", axis = 1).reset_index()
            .assign(Retorno = lambda _: (_.Resultado_Conjunto / (_.Payoff_Conjunto - _.Resultado_Conjunto) * 100).round(2).astype(str) + "%",
                    Payoff = lambda _: _.Payoff_Conjunto.round(2),
                    Resultado = lambda _: _.Resultado_Conjunto.round(2),
                    St = lambda _: _.St.round(2))
            .filter(["St", "Payoff", "Resultado", "Retorno"])
        )
        custo_estrategia = tabela.head(1).assign(custo = lambda _: _.Payoff - _.Resultado).custo

        header = (
            "-" * 50 + f"\n DADOS FINAIS DA ESTRATÉGIA {self.strategy} \n" + "-" * 50,
            f"Você realizou uma estratégia com {self.df.variable.nunique() - 2} operações",
            f"O custo de montar a estratégia foi {round(float(custo_estrategia), 2)}",
            "-" * 50 + f"\n TABELA DE VALORES PARA A ESTRATÉGIA {self.strategy}"
        )
        
        open(f"output/{self.strategy}.txt", "w", encoding = "utf-8").write("\n".join(header) + "\n" + str(tabela))

estrategias = [Strategy(file) for file in os.listdir("input")]

for estrategia in estrategias:
    estrategia.plot()
    estrategia.report()
    print(f"Estratégia {estrategia.strategy} finalizada")