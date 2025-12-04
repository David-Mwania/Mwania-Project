# RNA-seq Project

## Overview
This directory contains data and analysis files generated during the Intro to RNA-seq course.  
Date: [12.04.2025]

## Directory Structure
- **raw_data/**  
  Contains the original FASTQ files from the sequencing center. These files are unmodified and should remain unchanged throughout the analysis.

- **meta/**  
  Stores metadata describing the samples, including experimental conditions, replicates, and preparation details. Example: a CSV file listing sample names, treatments, and sequencing parameters.

- **logs/**  
  Holds log files that record commands run, software versions, parameters used, and summary statistics (e.g., number of reads trimmed, alignment rates). Useful for reproducibility and troubleshooting.

- **results/**  
  Contains outputs from analysis tools and workflows. Subdirectories can be created for each step (e.g., `results/fastqc`, `results/alignment`, `results/counts`).

- **scripts/**  
  Includes custom scripts (shell, Python, R) used to run analyses. Keeping scripts here ensures workflows can be repeated or shared easily.

## Notes
- Raw data should remain untouched.  
- Each analysis step should have its own subfolder in `results/`.  
- Metadata must be updated whenever new samples are added.  
- Logs should capture both successful runs and errors for transparency.  
