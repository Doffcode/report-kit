import matplotlib.pyplot as plt, os
plt.style.use(os.path.expanduser("~/report-kit/styles/report.mplstyle"))
rounds = list(range(1, 21))
fedavg = [0.42 + 0.5*(1-0.85**r) for r in rounds]
fedprox = [0.45 + 0.5*(1-0.82**r) for r in rounds]
fig, ax = plt.subplots()
ax.plot(rounds, fedavg, marker="o", ms=3, label="FedAvg")
ax.plot(rounds, fedprox, marker="s", ms=3, label="FedProx")
ax.set_xlabel("Communication round"); ax.set_ylabel("Test accuracy"); ax.legend()
for ext in ("svg", "png", "pdf"): fig.savefig(f"figures/accuracy.{ext}")
