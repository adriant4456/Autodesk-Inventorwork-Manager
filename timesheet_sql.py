import os
import sys
from sqlalchemy import Column, ForeignKey, Integer, String, DateTime
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from sqlalchemy import create_engine

Base = declarative_base()

class TimeEntry(base):
    __tablename__ = 'time'
    id = Column(Integer, primary_key = True)
    time = Column(Integer, nullable = False)
    date =  Column(DateTime, default =  datetime.datetime.utcnow)
    
