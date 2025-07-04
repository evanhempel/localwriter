This directory is intended to hold the vendored LiteLLM library and its dependencies for the localwriter extension.

To populate this directory, run the following commands in the root of the repository:
  mkdir -p lib
  pip install litellm -t lib

This ensures that all necessary dependencies are included for use within LibreOffice's Python environment, which may not have access to external packages.

Note for Developers:
- Ensure that the 'lib' directory is included when building the .oxt file for distribution.
- If you are testing locally, you may need to restart LibreOffice after adding or updating the contents of this directory to ensure the changes are recognized.
