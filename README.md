# ModernJsonInVBA

Deterministic JSON → Excel Tables  
Pure VBA. No dependencies. No silent schema drift.

---

## Introduction

**ModernJsonInVBA** is a single-module JSON engine built for serious Excel workflows.

It takes JSON text and materializes it into structured Excel tables (ListObjects) with fully predictable behavior.

This project focuses on control, stability, and repeatability.  
Every structural rule is explicit.  
Every schema change is intentional.  
Every failure is deterministic.

---

## What This Solves

Working with JSON in Excel often leads to:

- Columns appearing in different orders  
- Tables silently changing shape  
- Layout drift over time  
- Hidden dependencies on external libraries  
- Fragile refresh logic  

ModernJsonInVBA prevents those outcomes through:

- Stable column discovery order  
- Strict structural validation  
- Explicit schema controls  
- Deterministic error behavior  

When the JSON structure does not match the declared table root, execution stops with a clear, stable error. No guessing. No fallback tables.

Excel becomes a controlled data surface rather than a loosely interpreted one.

---

## Core Capabilities

- Parse JSON into VBA Variants (objects, arrays, primitives)  
- Convert VBA structures back into JSON  
- Flatten and rebuild object graphs  
- Discover array-of-object roots  
- Convert JSON tables into 2D arrays  
- Upsert Excel ListObjects deterministically  
- Enforce strict schema contracts when required  

All implemented in pure VBA.

No `Scripting.Dictionary`  
No COM references  
No external dependencies
