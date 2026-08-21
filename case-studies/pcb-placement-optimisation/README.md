# PCB Placement Optimisation

## Context

A team project modelled the assignment and operating sequence of three PCB placement machines. The workflow connects component requirements, feeder locations, placement-head compatibility, and machine movement into reproducible configuration and operation files.

## My contribution

As team lead, I focused on the attachment-assignment stage:

- designed and generated multiple candidate head-assignment strategies;
- compared coverage, workload balance, redundancy, and route-efficiency considerations;
- produced machine-readable assignment plans for downstream placement sequencing;
- helped standardise intermediate CSV schemas and validate integration across project stages.

The contribution is supported by my commit history in the original team repository. Other modules and the final report remain credited to the full group.

## Technical approach

```text
Raw component layout
        ↓
Standardised component / feeder / head mappings
        ↓
Candidate machine-head assignments
        ↓
Coverage and operational scoring
        ↓
Placement sequence generation and CSV validation
```

The repository uses relative paths and separates raw, intermediate, and output data. A validator checks required schemas and verifies that the placement logs cover all required actions.

## Evidence

- [Original team repository](https://github.com/zhloxy0907-bit/MATH-6119)
- GitHub contributor: [`goriyuki`](https://github.com/goriyuki)

## Notes

This page is an individual case-study summary, not a claim of sole authorship. Source code, data, reports, and teammate contributions remain in the original repository and its history.
