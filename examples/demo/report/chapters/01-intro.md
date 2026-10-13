# Introduction

Training machine learning models usually needs all data in one place. In many settings, such as hospitals or phones, data cannot leave the device. Federated learning solves this by sending the model to the data [@mcmahan2017fedavg].

## Problem Statement

When client data is non-IID, plain averaging of client updates becomes unstable. FedProx adds a proximal term to limit how far each client drifts [@li2020fedprox].

## Objectives

1. Build a federated training pipeline with a central aggregation server.
2. Compare FedAvg and FedProx on non-IID data.
3. Measure accuracy against communication rounds.
