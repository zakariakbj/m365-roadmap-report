# Changelog

All notable changes to this project are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html):

- **major**: a change that requires updating your configuration (`env.ini`, `workloads.psd1`, parameters)
- **minor**: a new feature, compatible with existing configurations
- **patch**: a bug fix

## [Unreleased]

## [1.0.0] - 2026-10-21

First public release.

### Added

- Weekly HTML email report of Microsoft 365 changes from the Message Center (Microsoft Graph) and the public Microsoft 365 Roadmap.
- Workloads configured in `workloads.psd1`: Message Center services and Roadmap products to match, with wildcards. Exchange, Outlook and Security (Defender for Office 365, Purview) by default.
- `-ListWorkloads` lists the Roadmap products and Message Center services available, with their number of items.
- `-OutFile` writes the report to a file instead of sending it, `-Days` sets the period, `-ConfigPath` and `-WorkloadsPath` point to other configuration files.
- Settings from environment variables or `env.ini`, environment variables first.
- Roadmap status badges (in development, rolling out, launched, cancelled) and a badge for each workload.
- Runs on PowerShell 7 (Windows, Linux, macOS) and Windows PowerShell 5.1, without external modules.

[Unreleased]: https://github.com/zakariakbj/m365-roadmap-report/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/zakariakbj/m365-roadmap-report/releases/tag/v1.0.0
