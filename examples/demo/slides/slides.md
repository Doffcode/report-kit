<!-- Demo deck: a project review. Each slide's design is chosen from what it says; the build lists the choices. -->

# Problem

## Raw patient data cannot leave the hospital

Hospitals hold the data a good model needs, but privacy rules keep every record on site.

- **Privacy**: records cannot be pooled in one place
- **Data**: each hospital sees only its own patients
- **Accuracy**: a model trained alone generalises poorly

::: notes
Start with the hospital example: records cannot be pooled, but a model can travel.
:::

## Clients do not share one data distribution

- **Non-IID data**: each client's data follows a different distribution, so one client's samples say little about another's

::: notes
Definition slide: the one term the rest of the talk depends on.
:::

## Objectives

1. Build a federated training pipeline with a central aggregation server
2. Compare FedAvg and FedProx on Independent and Identically Distributed (IID) and non-IID data
3. Measure accuracy against communication rounds

::: notes
Objective 2 is the core of the project; 1 and 3 support it.
:::

# Background

## Earlier work leaves non-IID stability open

| Paper | Method | Gap |
|:--|:--|:--|
| McMahan et al. [@mcmahan2017fedavg] | Average client models each round | Unstable on non-IID clients |
| Li et al. [@li2020fedprox] | Proximal term on the local objective | Not tested on small hospital splits |
| This work | FedProx on non-IID hospital splits | n/a |

::: notes
The gap column is highlighted: it is why this project exists.
:::

## FedProx keeps clients together where FedAvg lets them drift

### FedAvg (existing)

- Clients train freely between rounds
- Models drift apart on non-IID data
- Accuracy oscillates

### FedProx (proposed) {.win}

- A proximal term limits each client's drift
- Updates stay close to the global model
- Accuracy rises steadily

::: notes
Comparison: the proposed method is outlined; the existing one is quieter.
:::

# Method

## One server coordinates every client

- The server sends the global model each round
- Clients train on their own data
- Only weights travel back, never data

![System architecture.](figures/architecture)

::: notes
Diagram with its three notes numbered under it.
:::

## How one training round works

1. Server sends the global model
2. Each client trains a few local epochs
3. Clients send back their weights
4. Server averages them into the next model

::: notes
Process: shown as steps because the order matters.
:::

# Results

## FedProx ends 3 points higher after 20 rounds {.results}

- 0.92 :: FedProx accuracy
- 0.89 :: FedAvg accuracy
- 20 :: communication rounds

![Final test accuracy.](data/final.csv){type=bar x="Method" y="Accuracy" ymin=0.8 ymax=0.95}

::: notes
The numbers, and the chart they come from.
:::

## The gap opens within the first five rounds

![Test accuracy over 20 rounds; bands show the spread across 3 seeds.](data/accuracy.csv){x="Communication round" y="Test accuracy"}

FedProx is ahead by round 5 and never falls behind: the proximal term pays off early.

::: notes
Point at round 5.
:::

## FedProx helps, at a small cost

### Advantages

- No raw data leaves a client
- Only one extra parameter, μ
- Stable on non-IID splits

### Limitations

- μ must be tuned per dataset
- Slower rounds on weak devices

::: notes
Be upfront about the limits: reviewers ask.
:::

## Conclusion

- **Stable**: FedProx converges where FedAvg oscillates
- **Accurate**: 3 points higher after 20 rounds
- **Cheap**: one extra term, no extra communication

::: notes
Three takeaways; say them in this order.
:::

## Future work

- More clients :: scale to more hospitals
- Real data :: test on a hospital partner's records
- Privacy :: add differential privacy to the updates

::: notes
Roadmap: what comes next, in order.
:::

## References {.references}

## Why $\mu$ = 0.01? {.backup}

- Larger $\mu$ keeps clients close to the global model but slows learning
- 0.01 was the best of the values tried

::: notes
Backup for the viva question on hyper-parameters.
:::
