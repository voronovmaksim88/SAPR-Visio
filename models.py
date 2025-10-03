# models.py
"""
Модуль для работы с базой данных через SQLAlchemy
"""

from sqlalchemy import MetaData, Integer, String, ForeignKey, Date, Boolean, Text, DateTime, Table
from sqlalchemy.ext.asyncio import AsyncAttrs
from sqlalchemy.orm import validates
from sqlalchemy.orm import DeclarativeBase
from sqlalchemy.orm import Mapped  # используется для аннотации типов столбцов.
from sqlalchemy.orm import mapped_column  # функция для определения столбцов с дополнительными
from sqlalchemy.orm import relationship  # используется для связи таблиц
from sqlalchemy import Interval  # Импортируем Interval для работы с временными интервалами
import uuid
from uuid import uuid4
from sqlalchemy import Column
from sqlalchemy.dialects.postgresql import UUID

from typing import Optional
from typing import List

from datetime import date, UTC
from datetime import datetime
from datetime import timedelta

# Переменная, которая хранит информацию о таблицах
metadata = MetaData()


# Класс базы данных. Используется для создания таблиц в базе данных
class Base(DeclarativeBase):
    pass


class Country(Base):
    """Страны"""
    __tablename__ = 'countries'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)

    def __str__(self) -> str:
        return self.name


class Manufacturer(Base):
    """Производители, это не конкретное юрлицо это типа брэнд"""
    __tablename__ = 'manufacturers'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)
    country_id: Mapped[int] = mapped_column(ForeignKey('countries.id'), nullable=False)

    # Опционально: добавление отношения к Country
    country: Mapped["Country"] = relationship()
    equipments: Mapped[List["Equipment"]] = relationship(back_populates="manufacturer")

    def __repr__(self) -> str:
        return f"Manufacturer(id={self.id}, name={self.name})"


class Currency(Base):
    """Класс "Валюты". Тут просто имена валют"""
    __tablename__ = 'currencies'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(3), nullable=False, unique=True)

    # Отношения
    equipments: Mapped[List["Equipment"]] = relationship(back_populates="currency")

    def __repr__(self) -> str:
        return f"Currency(id={self.id!r}, name={self.name!r})"


class City(Base):
    """Города"""
    __tablename__ = 'cities'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)
    country_id: Mapped[int] = mapped_column(ForeignKey('countries.id'), nullable=False)

    country: Mapped["Country"] = relationship()
    counterparties: Mapped[List["Counterparty"]] = relationship(back_populates="city")

    def __repr__(self) -> str:
        return f"City(id={self.id!r}, name={self.name!r})"


class CounterpartyForm(Base):
    """Формы контрагентов"""
    __tablename__ = 'counterparty_form'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)

    # Добавляем обратную связь, если нужно будет получать всех контрагентов этой формы
    counterparties: Mapped[List["Counterparty"]] = relationship(back_populates="form")

    def __repr__(self) -> str:
        return f"CounterpartyForm(id={self.id!r}, name={self.name!r})"


class Counterparty(Base):
    """ Контрагенты, юрлица, ИП, ЧЛ и т.д все с кем мы сотрудничаем"""
    __tablename__ = 'counterparty'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)
    note: Mapped[str | None] = mapped_column(String, nullable=True)
    city_id: Mapped[int | None] = mapped_column(ForeignKey('cities.id'), nullable=True)
    form_id: Mapped[int] = mapped_column(ForeignKey('counterparty_form.id'), nullable=False)
    orders: Mapped[List["Order"]] = relationship(back_populates="customer")

    form: Mapped["CounterpartyForm"] = relationship(back_populates="counterparties")
    city: Mapped[Optional["City"]] = relationship(back_populates="counterparties")
    # Можно использовать lazy='joined' или lazy='selectin' здесь,
    # но лучше управлять загрузкой в самом запросе через options()

    def __repr__(self) -> str:
        return f"Counterparty(id={self.id!r}, name={self.name!r})"



class EquipmentType(Base):
    """Типы оборудования"""
    __tablename__ = 'equipment_types'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)
    description: Mapped[str | None] = mapped_column(String, nullable=True)

    # Отношения
    equipments: Mapped[List["Equipment"]] = relationship(back_populates="type")

    def __repr__(self) -> str:
        return f"EquipmentType(id={self.id}, name={self.name})"


class Equipment(Base):
    """
    Класс "Оборудование"
    """
    __tablename__ = 'equipment'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False)  # Имя
    model: Mapped[Optional[str]] = mapped_column(String(64), nullable=True, unique=True)  # Модель

    # Артикул, код поставщика
    vendor_code: Mapped[Optional[str]] = mapped_column(String(32), nullable=True, unique=True)
    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)  # Описание
    type_id: Mapped[Optional[int]] = mapped_column(ForeignKey('equipment_types.id'), nullable=True)  # Тип оборудования

    # Производитель
    manufacturer_id: Mapped[Optional[int]] = mapped_column(ForeignKey('manufacturers.id'), nullable=True)
    price: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)  # Цена
    currency_id: Mapped[Optional[int]] = mapped_column(ForeignKey('currencies.id'), nullable=True)  # Валюта
    relevance: Mapped[bool] = mapped_column(Boolean, default=True)  # Актуальность
    price_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)  # Дата обновления цены

    # пока не решил как хранить фотки
    # photo: Mapped[Optional[str]] = mapped_column(String, nullable=True)  # Путь к фото

    # Отношения
    type: Mapped["EquipmentType"] = relationship(back_populates="equipments")
    manufacturer: Mapped["Manufacturer"] = relationship(back_populates="equipments")
    currency: Mapped["Currency"] = relationship(back_populates="equipments")

    # Для наследования с одной таблицей
    discriminator = mapped_column(String(50))
    __mapper_args__ = {
        'polymorphic_on': discriminator,
        'polymorphic_identity': 'equipment'
    }

    def __repr__(self) -> str:
        return f"Equipment(id={self.id!r}, name={self.name!r}, model={self.model!r})"


class ControlCabinetMaterial(Base):
    __tablename__ = 'control_cabinet_materials'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False)

    # Отношения
    control_cabinets: Mapped[List["ControlCabinet"]] = relationship(back_populates="material")

    def __str__(self) -> str:
        return self.name


class Height(Base):
    """Таблица высот шкафов"""
    __tablename__ = 'heights'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    value: Mapped[int] = mapped_column(Integer, nullable=False)

    # Отношения
    control_cabinets: Mapped[List["ControlCabinet"]] = relationship(back_populates="height_ref")

    def __repr__(self) -> str:
        return f"Height(id={self.id!r}, value={self.value!r})"


class Width(Base):
    """Таблица ширин шкафов"""
    __tablename__ = 'widths'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    value: Mapped[int] = mapped_column(Integer, nullable=False)

    # Отношения
    control_cabinets: Mapped[List["ControlCabinet"]] = relationship(back_populates="width_ref")

    def __repr__(self) -> str:
        return f"Width(id={self.id!r}, value={self.value!r})"


class Depth(Base):
    """Таблица глубин шкафов"""
    __tablename__ = 'depths'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    value: Mapped[int] = mapped_column(Integer, nullable=False)

    # Отношения
    control_cabinets: Mapped[List["ControlCabinet"]] = relationship(back_populates="depth_ref")

    def __repr__(self) -> str:
        return f"Depth(id={self.id!r}, value={self.value!r})"


# Класс для степеней защиты по ip, возможно что он не только для корпусов шкафов пригодится
class Ip(Base):
    """
    IP - степень защиты шкафов автоматики
    """
    __tablename__ = 'ips'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False)

    # Отношения
    control_cabinets: Mapped[List["ControlCabinet"]] = relationship(back_populates="ip")


class ControlCabinet(Equipment):
    """
    Корпуса шкафов автоматики
    """
    __tablename__ = 'control_cabinets'
    id = Column(Integer, ForeignKey('equipment.id'), primary_key=True)

    __mapper_args__ = {
        'polymorphic_identity': 'control_cabinet'
    }

    material_id: Mapped[int] = mapped_column(ForeignKey('control_cabinet_materials.id'), nullable=False)
    ip_id: Mapped[int] = mapped_column(ForeignKey('ips.id'), nullable=False)

    # Новые поля для связи с таблицами размеров
    height_id: Mapped[int | None] = mapped_column(ForeignKey('heights.id'), nullable=True)
    width_id: Mapped[int | None] = mapped_column(ForeignKey('widths.id'), nullable=True)
    depth_id: Mapped[int | None] = mapped_column(ForeignKey('depths.id'), nullable=True)

    # Отношения
    material: Mapped["ControlCabinetMaterial"] = relationship(back_populates="control_cabinets")
    ip: Mapped["Ip"] = relationship(back_populates="control_cabinets")
    
    # Новые отношения с таблицами размеров
    height_ref: Mapped[Optional["Height"]] = relationship(back_populates="control_cabinets", foreign_keys=[height_id])
    width_ref: Mapped[Optional["Width"]] = relationship(back_populates="control_cabinets", foreign_keys=[width_id])
    depth_ref: Mapped[Optional["Depth"]] = relationship(back_populates="control_cabinets", foreign_keys=[depth_id])



class SensorMeasuredValues(Base):
    """
    Таблица для хранения измеряемых значений датчиков
    """
    __tablename__ = "sensor_measured_values"
    
    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False, unique=True, index=True)
    
    # Отношения определены через backref в классе Sensor
    
    def __str__(self) -> str:
        return self.name


class SensorTypes(Base):
    """
    Таблица для хранения типов датчиков
    """
    __tablename__ = "sensor_types"
    
    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False, unique=True, index=True)
    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    
    # Отношения определены через backref в классе Sensor
    
    def __str__(self) -> str:
        return self.name


class SensorsShapeType(Base):
    """
    Таблица для хранения типов отображения датчиков на схемах Visio
    """
    __tablename__ = "sensors_shape_type"
    
    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False, unique=True, index=True)
    description: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    shape_code: Mapped[Optional[str]] = mapped_column(String(32), nullable=True)
    
    # Отношения
    sensors: Mapped[List["Sensor"]] = relationship(back_populates="shape_type")
    
    def __str__(self) -> str:
        return self.name


# Промежуточная таблица для связи многие-ко-многим между Sensor и SensorMeasuredValues
sensor_measured_values_association = Table(
    'sensor_measured_values_association',
    Base.metadata,
    Column('sensor_id', Integer, ForeignKey('sensors.id'), primary_key=True),
    Column('measured_value_id', Integer, ForeignKey('sensor_measured_values.id'), primary_key=True)
)


# Промежуточная таблица для связи многие-ко-многим между Sensor и SensorTypes
sensor_types_association = Table(
    'sensor_types_association',
    Base.metadata,
    Column('sensor_id', Integer, ForeignKey('sensors.id'), primary_key=True),
    Column('sensor_type_id', Integer, ForeignKey('sensor_types.id'), primary_key=True)
)


class Sensor(Equipment):
    """
    Датчики
    """
    __tablename__ = 'sensors'
    id = Column(Integer, ForeignKey('equipment.id'), primary_key=True)

    __mapper_args__ = {
        'polymorphic_identity': 'sensor'
    }

    # Связь с типом отображения (один-ко-многим)
    sensors_shape_type_id: Mapped[Optional[int]] = mapped_column(ForeignKey('sensors_shape_type.id'), nullable=True)
    
    # Отношения
    shape_type: Mapped[Optional["SensorsShapeType"]] = relationship(back_populates="sensors")
    
    # Связи многие-ко-многим
    measured_values: Mapped[List["SensorMeasuredValues"]] = relationship(
        secondary=sensor_measured_values_association,
        backref="sensors"
    )
    
    sensor_types: Mapped[List["SensorTypes"]] = relationship(
        secondary=sensor_types_association,
        backref="sensors"
    )