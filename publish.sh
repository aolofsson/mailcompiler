##############################################################################
# Publish to pypi
##############################################################################

# 1. Create and activate a virtualenv
python3 -m venv venv
source venv/bin/activate

# 2. Install the publishing tools (build = makes the package, twine = uploads it)
pip install --upgrade build twine

# 3. Clean any old artifacts, then build the wheel + sdist into dist/
rm -rf dist build mailcompiler.egg-info
python -m build

# 4. Sanity-check the built metadata
twine check dist/*

# 5. (optional) smoke-test on TestPyPI first (separate account/auth!!)
#twine upload --repository testpypi dist/*
#    username: __token__   password: <your TestPyPI token>

# 6. Publish to real PyPI
twine upload dist/*
#    username: __token__   password: <your pypi.org token>
