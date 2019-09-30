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
    entered_date =  Column(DateTime, default =  datetime.datetime.utcnow)
    machine_id = Column(Integer, ForeignKey('machine.id'))

class Machine(base):
    __tablename__ = 'machine'
    id = Column(Integer, primary_key = True)
    network = Column(Integer, nullable = False)
    activity = Column(Integer, nullable = False)
    name = Column(String(250), nullable = False)
    entries = relationship("TimeEntry")
    
