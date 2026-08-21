# Healthcare Appointment Simulation

## Question

How can appointment capacity be allocated to reduce urgent-patient waiting time while keeping overall average waiting time within five days?

## Model

I built a discrete-event simulation in AnyLogic with two patient classes—urgent and routine—sharing limited daily appointment capacity. The model compared a first-available-slot baseline with dynamic-delay policies that respond to the current urgent workload.

## Experiment

- repeated each candidate policy across 200 Monte Carlo replications;
- compared urgent and overall mean waiting times;
- used confidence intervals to distinguish systematic improvements from simulation noise;
- checked the operational constraint on overall mean waiting time before recommending a policy.

## Result

The selected policy reduced urgent mean waiting time by approximately **49%** relative to the baseline while keeping overall mean waiting time below **five days** in the evaluated scenario.

## Publication note

This repository contains only a concise case summary while the degree programme is active. The assessment brief, submitted report, model file, and course-specific parameter set are withheld. A walkthrough can be provided privately for recruitment where permitted.
