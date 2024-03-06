#!/usr/bin/env python
# coding: utf-8

# In[2]:


from setuptools import setup, find_packages

setup(
    name='Enso_Package',
    version='0.1.0',
    packages=find_packages(),
    install_requires=[
        # Hier Ihre Abhängigkeiten, wie in requirements.txt
        'pandas>=1.3.0',
        'numpy>=1.21.0',
        'Office365-REST-Python-Client>=2.3.5'
    ],
    author= 'Philipp Karlson',
    author_email= 'philipp.karlson@myenso.de',
    url='https://github.com/philippkarlson-myenso/test.git',
)


# In[ ]:




