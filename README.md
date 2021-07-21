# Agile Flex

## Prerequisites

- [Python 3.7]
- See the requirements.txt file

## Overall Project Description

Harmonic Spectrums are injected across a test feeder in OpenDSS. A load flow is carried out then harmonic analysis.
There are scripts to carry out single snapshot for individual and mixtures of EVs. Then Monte-Carlo of randomly assigned EVs across feeder.

## Scripts:

dss_run.py - carries out monte-carlo studies of harmonics for random EVs on test feeder. Results are returned as .pickle files into the results folder
Harmonic_plots.py - plots the results from dss_run.py script
Harmonic_Cancellation - Looks at cancellation from different EV types on same phase compared to same EV type
Harmonic_DeratedComparison - Comparing harmonics from different derated capacities of same type
Harmonic_DeratedComparisonDiversity - Comparing harmonics from different derated capacities of different types
Harmonic_SingleBus - Comparing harmonics for each EV type for CC and CV
Harmonic_SingleBus - Comparing harmonics for each EV type for derated capacity

