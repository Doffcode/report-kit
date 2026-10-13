# Results

## Experimental Setup

@tbl-setup lists the training configuration. The header row uses merged cells.

+----------------------+---------------------+------------+
| Method               | Training            | Accuracy   |
|                      +----------+----------+            |
|                      | Rounds   | Epochs   |            |
+======================+==========+==========+============+
| FedAvg               | 20       | 5        | 0.89       |
+----------------------+----------+----------+------------+
| FedProx ($\mu=0.01$) | 20       | 5        | 0.92       |
+----------------------+----------+----------+------------+

: Training configuration and final accuracy. {#tbl-setup}

## Accuracy

@fig-accuracy shows test accuracy over 20 rounds. FedProx converges faster and ends higher.

![Test accuracy against communication round.](figures/accuracy){#fig-accuracy width=80%}

As @tbl-setup shows, FedProx improves final accuracy by three points over FedAvg [@li2020fedprox].
