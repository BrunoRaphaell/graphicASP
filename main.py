# importing the libraries
import pandas as pd
import matplotlib.pyplot as plt

archive = 'gd_barra_{}.xlsx'  # Opening the file
pq_bars = [4, 5, 7, 10, 11, *[i for i in range(13, 58)]]  # Setting the PQ bars

# fixed variables for the loses
losses_without_gd_MW = 152.3
losses_without_gd_MVAr = 502.8
losses_with_split_GD_MW = 40.8
losses_with_split_GD_MVar = 31.5

# Variables used to store losses.
losses_MW = []
losses_MVAr = []
bars = []

# handling errors in case the file cannot be found
for bar in pq_bars:
    try:
        df = pd.read_excel(archive.format(bar), sheet_name='Tabela Resumo')
        bars.append(bar)
        # [0] -> Perdas MW
        # [1] -> Perdas em Mvar
        losses_MW.append(df['Perdas MW / Mvar'][0])
        losses_MVAr.append(df['Perdas MW / Mvar'][1])
    except FileNotFoundError as e:
        pass

fig = plt.figure(1)

# Plot perdas(MW)
ax1 = fig.add_subplot(211)  # Add an Axes to the current figure or retrieve an existing Axes.
ax1.grid(True, axis='y', linestyle='dashed')  # Configure the grid lines.
ax1.set_axisbelow(True)  # Get whether axis ticks and gridlines are above or below most artists.
plt.title("Perdas Ativas com GD")  # define the title of the chart
ax1.plot([1, 57], [losses_without_gd_MW, losses_without_gd_MW], 'r--')  # marking line for losses without distributed generation
ax1.plot([1, 57], [losses_with_split_GD_MW, losses_with_split_GD_MW], 'g--')  # marker line for split non-generation losses
ax1.bar(bars, losses_MW)
ax1.legend(
    ['Perdas Ativas no Sistema sem GD',
     'Perdas Ativas no Sistema com GD Dividida',
     'Perdas Ativas no Sistema com GD Localizada'])
ax1.set_xticks(pq_bars)
ax1.set_ylabel('Potência Ativa (MW)')
ax1.set_xlabel('Barra')

# Plot perdas reativas(MVar)
ax2 = fig.add_subplot(212)
ax2.grid(True, axis='y', linestyle='dashed')
ax2.set_axisbelow(True)
plt.title("Perdas Reativas com GD")
ax2.plot([1, 57], [losses_without_gd_MVAr, losses_without_gd_MVAr], 'r--')
ax2.plot([1, 57], [losses_with_split_GD_MVar, losses_with_split_GD_MVar], 'g--')
ax2.bar(bars, losses_MVAr)
ax2.legend(
    ['Perdas Reativas no Sistema sem GD',
     'Perdas Reativas no Sistema com GD Dividida',
     'Perdas Reativas no Sistema com GD Localizada'])
ax2.set_xticks(pq_bars)
ax2.set_ylabel('Potência Reativa (MVAr)')
ax2.set_xlabel('Barra')

plt.show()
