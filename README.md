# Yaoer Drone Trade Workbench

This repository is a working folder for a drone export startup pipeline. It combines:

- spreadsheet artifacts for market research and supplier development
- generated PNG snapshots for quick review
- builder scripts used to regenerate the workbooks
- contact copy used for supplier outreach

The current focus is on drone accessories, FPV products, agricultural drones, and industrial UAV supply-chain development.

## Repository Layout

```text
outputs/
  contact_queue/
  drone_trade_startup/
  guangdong_supplier_pool/
  supplier_pool_restart/
tools/
  drone_trade_builder/
  guangdong_supplier_builder/
  supplier_pool_builder/
  supplier_pool_restart/
```

## Outputs

### `outputs/drone_trade_startup`

Main startup workbook and preview images for the 90-day drone export plan.

- `无人机外贸启动总表.xlsx`: operating workbook
- `总览Dashboard.png`: dashboard preview
- `产品线SKU.png`: product-line preview
- `目标市场.png`: target-market preview
- `潜在客户.png`: lead-tracking preview
- `报价&利润.png`: quote and margin preview
- `90天任务.png`: execution plan preview
- `verification.txt`: artifact verification record

### `outputs/guangdong_supplier_pool`

Filtered Guangdong supplier pool for small-batch-first sourcing.

- `广东无人机及配件工厂_小批量优先版.xlsx`: supplier workbook
- `广东工厂池.png`: main pool preview
- `筛选口径.png`: filtering criteria
- `联系核验问题.png`: supplier verification prompts
- `暂不优先.png`: lower-priority list preview
- `verification.txt`: artifact verification record

### `outputs/supplier_pool_restart`

Restarted supplier sourcing pack with a rebuilt supplier pool and outreach material.

- `无人机供应商开发池_重新收集版.xlsx`: rebuilt supplier workbook
- `供应商池.png`: supplier pool preview
- `来源说明.png`: source-notes preview
- `询价话术.png`: inquiry copy preview
- `验厂问题.png`: factory-check questions preview
- `verification.txt`: artifact verification record

### `outputs/contact_queue`

- `非玩具无人机供应商打招呼队列.md`: ready-to-send supplier outreach queue

## Builder Scripts

Each builder script generates one workbook package under `outputs/`.

- `tools/drone_trade_builder/build.mjs`
  Builds the startup workbook for products, markets, leads, quotes, suppliers, content, and 90-day execution tracking.
- `tools/guangdong_supplier_builder/build.mjs`
  Builds the Guangdong supplier screening workbook.
- `tools/supplier_pool_builder/build.mjs`
  Builds the original supplier pool workbook.
- `tools/supplier_pool_restart/build.mjs`
  Builds the restarted supplier development workbook and related materials.

These scripts use `@oai/artifact-tool` and write directly to the absolute output paths currently configured inside each file.

## Current Workflow

1. Update source assumptions in the builder scripts.
2. Rebuild the workbook package you need.
3. Review the generated `.xlsx`, `.png`, and `verification.txt` outputs.
4. Use the contact queue and supplier pool to run outreach and follow-up.
5. Commit updated outputs once the new version is verified.

## Notes

- The repository currently stores generated artifacts alongside source scripts on purpose.
- `gh` and regular `git` are both configured locally now, so future pushes and GitHub CLI operations can run from this repo.
- The builder scripts currently depend on absolute local paths, so moving the repo will require updating those output paths.
