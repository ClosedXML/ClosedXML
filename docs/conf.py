# Configuration file for the Sphinx documentation builder.
#
# This file only contains a selection of the most common options. For a full
# list see the documentation:
# http://www.sphinx-doc.org/en/master/config

# -- Path setup --------------------------------------------------------------

# If extensions (or modules to document with autodoc) are in another directory,
# add these directories to sys.path here. If the directory is relative to the
# documentation root, use os.path.abspath to make it absolute, like shown here.
#
import os
import sys
import sphinx_rtd_theme
sys.path.insert(0, os.path.abspath('../ClosedXML'))

# -- Project information -----------------------------------------------------

project = 'ClosedXML'

# The full version, including alpha/beta/rc tags
release = '0.102.0'

# -- General configuration ---------------------------------------------------

# Add any Sphinx extension module names here, as strings. They can be
# extensions coming with Sphinx (named 'sphinx.ext.*') or your custom
# ones.
extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.todo',
    'sphinx_rtd_theme',
    'sphinx_csharp',
    'breathe',
    'linuxdoc.rstFlatTable', # Table with row/col spans
]

# Add any paths that contain templates here, relative to this directory.
templates_path = ['_templates']

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = ['_build', 'Thumbs.db', '.DS_Store']

# -- Options for doxygen output ----------------------------------------------
import subprocess, os

read_the_docs_build = os.environ.get('READTHEDOCS', None) == 'True'

if read_the_docs_build:
     subprocess.call('doxygen', shell=True)
     subprocess.call('pip install git+https://github.com/rogerbarton/sphinx-csharp.git', shell=True)

# -- Options for HTML output -------------------------------------------------

# The theme to use for HTML and HTML Help pages.  See the documentation for
# a list of builtin themes.
#
html_theme = 'sphinx_rtd_theme'
pygments_style = 'sphinx'

# Add any paths that contain custom static files (such as style sheets) here,
# relative to this directory. They are copied after the builtin static files,
# so a file named "default.css" will overwrite the builtin "default.css".
html_static_path = []
html_css_files = []
html_theme_options = {
    'navigation_depth': 2,
}

# Breathe Configuration
breathe_projects = { "ClosedXML" : "./xml" }
breathe_default_project = "ClosedXML"
