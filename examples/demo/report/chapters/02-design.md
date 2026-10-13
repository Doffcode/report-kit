# System Design

The system has one aggregation server and $N$ clients, shown in @fig-architecture. Each client trains locally and sends only model weights.

![High-level architecture of the federated learning system.](figures/architecture){#fig-architecture width=85%}

The training loop on each client is shown in @fig-flow.

![Client-side training flow.](figures/flow){#fig-flow width=45%}

## Aggregation

The server combines client weights with a weighted average, given in @eq-fedavg:

$$
w_{t+1} = \sum_{k=1}^{K} \frac{n_k}{n} \, w_{t+1}^{k}
$$ {#eq-fedavg}

FedProx changes each client's local objective by adding a proximal term (@eq-fedprox):

$$
\min_{w} \; F_k(w) + \frac{\mu}{2} \|w - w_t\|^2
$$ {#eq-fedprox}

## Algorithm

::: {#alg-fedprox}


1. Server sends global weights $w_t$ to selected clients.
2. **for** each client $k$ in parallel **do**
3. $\quad$ $w^k \leftarrow$ minimise $F_k(w) + \frac{\mu}{2}\lVert w - w_t\rVert^2$ for $E$ local epochs
4. $\quad$ send $w^k$ to server
5. **end for**
6. Server aggregates using @eq-fedavg.

FedProx training round.
:::
