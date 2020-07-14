"""
Facebook AI powered chatbot in Excel.

There are two functions:
  - parlai_create_world creates a world containing two agents (the AI and the human).
  - parlai_speak takes an input from the human and runs the model to get the AI response.

The entire conversation is returned by parlai_speak so it can be viewed in Excel
rather than just the last response.

See https://parl.ai/projects/blender/

This code is adapted from a previous video tutorial that can be found here:
https://www.youtube.com/watch?v=sKmII0VVLKk
"""
from .parlai_excel import *
