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


class Person(Base):
    """Люди - сотрудники, представители заказчиков и т.д."""
    __tablename__ = 'people'

    uuid: Mapped[UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, unique=True, nullable=False, default=uuid4)
    name: Mapped[str] = mapped_column(String, nullable=False)
    patronymic: Mapped[str | None] = mapped_column(String, nullable=True)  # Отчество
    surname: Mapped[str] = mapped_column(String, nullable=False)  # Фамилия
    phone: Mapped[str | None] = mapped_column(String, nullable=True)  # Телефон
    email: Mapped[str | None] = mapped_column(String, nullable=True)  # Email
    counterparty_id: Mapped[int | None] = mapped_column(ForeignKey('counterparty.id'), nullable=True)
    # Человек может не иметь принадлежности ни к одной компании
    birth_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    active: Mapped[bool] = mapped_column(Boolean, default=True)
    # Активен, если это сотрудник, то с ним можно работать в настоящий момент,
    # если это представитель заказчика, то он жив и ещё работает в нашей сфере
    note: Mapped[str | None] = mapped_column(Text, nullable=True)  # Примечание
    can_be_scheme_developer: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_be_assembler: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_be_programmer: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    can_be_tester: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)

    # Внешний ключ на User
    user_id: Mapped[int] = mapped_column(
        ForeignKey('users.id'),
        nullable=True,
        unique=True  # Обеспечивает уникальность связи
    )

    # Связь один-к-одному с User
    user: Mapped["User"] = relationship(
        back_populates="person",
        uselist=False
    )

    # Добавляем отношение к задачам
    tasks: Mapped[List["Task"]] = relationship(
        back_populates="executor",
        foreign_keys="[Task.executor_uuid]"
    )


    def __repr__(self) -> str:
        return f"Person(id={self.uuid!r}, name={self.name!r}, surname={self.surname!r})"

    # relations
    developed_boxes = relationship("BoxAccounting", back_populates="scheme_developer",
                                   foreign_keys="[BoxAccounting.scheme_developer_id]")
    assembled_boxes = relationship("BoxAccounting", back_populates="assembler",
                                   foreign_keys="[BoxAccounting.assembler_id]")
    programmed_boxes = relationship("BoxAccounting", back_populates="programmer",
                                    foreign_keys="[BoxAccounting.programmer_id]")
    tested_boxes = relationship("BoxAccounting", back_populates="tester",
                                foreign_keys="[BoxAccounting.tester_id]")
    timing_records = relationship("Timing", back_populates="executor",
                                  foreign_keys="[Timing.executor_id]")




# Вспомогательная таблица для связи многие-ко-многим между Order и Work
order_work = Table(
    'orders_works',
    Base.metadata,
    Column('order_serial', ForeignKey('orders.serial'), primary_key=True),
    Column('work_id', ForeignKey('works.id'), primary_key=True)
)


class Work(Base):
    """Работы, выполняемые по заказам"""
    __tablename__ = 'works'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)  # Название
    description: Mapped[str | None] = mapped_column(String, nullable=True)  # Описание
    active: Mapped[bool] = mapped_column(Boolean, default=True)
    # Активно, значит эту работу можно назначить новому проекту
    orders: Mapped[List["Order"]] = relationship(secondary="orders_works", back_populates="works")

    def __repr__(self) -> str:
        return f"Work(id={self.id!r}, name={self.name!r})"


class OrderStatus(Base):
    """
    Статусы заказов
    1 = "Не определён"
    2 = "На согласовании"
    3 = "В работе"
    4 = "Просрочено"
    5 = "Выполнено в срок"
    6 = "Выполнено НЕ в срок"
    7 = "Не согласовано"
    8 = "На паузе"
    """
    __tablename__ = 'order_statuses'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String, nullable=False, unique=True)  # Название
    description: Mapped[str | None] = mapped_column(String, nullable=True)  # Описание
    orders: Mapped[List["Order"]] = relationship(back_populates="status")

    def __repr__(self) -> str:
        return f"OrderStatus(id={self.id!r}, name={self.name!r})"


class Order(Base):
    """Таблица заказов (заявок, проектов)"""
    __tablename__ = 'orders'

    serial: Mapped[str] = mapped_column(String(16), primary_key=True)
    # Серийный номер заказа, формат NNN-MM-YYYY
    # NNN - порядковый номер в этом году
    # MM - месяц создания
    # YYYY - год создания
    name: Mapped[str] = mapped_column(String(128), nullable=False)  # Название
    customer_id: Mapped[int | None] = mapped_column(ForeignKey('counterparty.id'), nullable=True)  # id заказчика
    customer: Mapped["Counterparty"] = relationship(back_populates="orders", foreign_keys=[customer_id])
    priority: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Приоритет от 1 до 10
    status_id: Mapped[int] = mapped_column(ForeignKey('order_statuses.id'), nullable=False)  # Статус заказа
    status: Mapped["OrderStatus"] = relationship(back_populates="orders")
    start_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)  # Дата и время создания
    deadline_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)  # Дата и время дедлайна
    end_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)  # Дата и время окончания

    # relations
    boxes: Mapped[List["BoxAccounting"]] = relationship(back_populates="order")
    works: Mapped[List["Work"]] = relationship(secondary="orders_works", back_populates="orders")
    tasks: Mapped[List["Task"]] = relationship(back_populates="order")  # Отношение с Task
    comments: Mapped[List["OrderComment"]] = relationship(back_populates="order")
    timings: Mapped[List["Timing"]] = relationship(back_populates="order")  # Добавьте эту строку

    """
    Строки ниже нужны для грубой финансовой аналитики
    При создании заказа вписываем сколько надо под него денег заложить на товары, материалы и работы
    Потом в аналитике видно сколько денег в моменте надо под закрытие текущих заказов,
    далее смотрим сколько есть на счёте и делаем выводы
    После того как что-то оплачено из этого полностью, выставляем флаг что оплачено.
    Такой подход не даёт точного анализа расходов, но позволяет быстро определить текущую потребность в деньгах.
    Так же фиксируем сколько клиент денег должен ещё нам. 
    """
    materials_cost: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Стоимость материалов плановая
    materials_cost_fact: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Стоимость материалов плановая
    materials_paid: Mapped[bool] = mapped_column(Boolean, default=False)  # Материалы оплачены

    products_cost: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Стоимость товаров плановая
    products_cost_fact: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Стоимость товаров плановая
    products_paid: Mapped[bool] = mapped_column(Boolean, default=False)  # Товары оплачены

    work_cost: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Стоимость работ плановая
    work_cost_fact: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Стоимость работ плановая
    work_paid: Mapped[bool] = mapped_column(Boolean, default=False)  # Работы оплачены

    debt: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Задолженность нам план
    debt_fact: Mapped[int | None] = mapped_column(Integer, nullable=True)  # Задолженность нам уже вернули
    debt_paid: Mapped[bool] = mapped_column(Boolean, default=False)  # Задолженность оплачена

    @validates('priority')
    def validate_priority(self, key, value):  # noqa
        if value is not None:
            if value < 1 or value > 10:
                raise ValueError("Priority must be between 1 and 10")
        return value

    @validates('status_id')
    def validate_status_id(self, key, value):  # noqa
        if value < 1 or value > 8:
            raise ValueError("Status ID must be between 1 and 8")
        return value

    def __repr__(self) -> str:
        return f"Order(serial={self.serial!r}, name={self.name!r})"


class BoxAccounting(Base):
    """Таблица учёта шкафов """
    __tablename__ = 'box_accounting'
    serial_num: Mapped[int] = mapped_column(primary_key=True, unique=True)
    name: Mapped[str] = mapped_column(String(64), nullable=False)  # Название шкафа
    order_id: Mapped[str] = mapped_column(ForeignKey('orders.serial'), nullable=False)  # Заказ
    # Разработчик схемы
    scheme_developer_id: Mapped[uuid.UUID] = mapped_column(ForeignKey('people.uuid'), nullable=False)
    assembler_id: Mapped[uuid.UUID] = mapped_column(ForeignKey('people.uuid'), nullable=False)  # Сборщик
    programmer_id: Mapped[uuid.UUID | None] = mapped_column(ForeignKey('people.uuid'), nullable=True)  # Программист
    tester_id: Mapped[uuid.UUID] = mapped_column(ForeignKey('people.uuid'), nullable=False)  # Тестировщик

    # Определяем отношения. Пока не знаю зачем
    order = relationship("Order", back_populates="boxes")
    scheme_developer = relationship("Person", foreign_keys="[BoxAccounting.scheme_developer_id]",
                                    back_populates="developed_boxes")
    assembler = relationship("Person", foreign_keys="[BoxAccounting.assembler_id]", back_populates="assembled_boxes")
    programmer = relationship("Person", foreign_keys="[BoxAccounting.programmer_id]", back_populates="programmed_boxes")
    tester = relationship("Person", foreign_keys="[BoxAccounting.tester_id]", back_populates="tested_boxes")


class OrderComment(Base):
    """Таблица комментариев к заказам """
    __tablename__ = 'comments_on_orders'
    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    order_id: Mapped[str] = mapped_column(ForeignKey('orders.serial'), nullable=False)  # Заказ
    moment_of_creation: Mapped[Optional[datetime]] = mapped_column(DateTime, default=datetime.now,
                                                                   nullable=True)  # Дата и время публикации комментария
    text: Mapped[str] = mapped_column(Text, nullable=False)  # Текст комментария
    person_uuid: Mapped[uuid.UUID] = mapped_column(ForeignKey('people.uuid'), nullable=False)  # Автор комментария

    # relations
    order: Mapped["Order"] = relationship(back_populates="comments")


class TaskStatus(Base):
    """
    Статусы задач
    1 = "Не начата"
    2 = "В работе"
    3 = "На паузе"
    4 = "Завершена"
    5 = "Отменена"
    else = "?"
    """
    __tablename__ = 'task_statuses'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(16), nullable=False, unique=True)  # Название статуса задачи

    # Связь с задачами (если будет создана модель Task)
    tasks: Mapped[List["Task"]] = relationship(back_populates="status")

    def __repr__(self) -> str:
        return f"TaskStatus(id={self.id!r}, name={self.name!r})"


class TaskPaymentStatus(Base):
    """
    Статусы оплаты за задачу
    1 = "Нет оплаты", задача не предполагает оплату
    2 = "Возможна", задача в работе если исполнитель сделает её вовремя и качественно, то получит оплату
    3 = "Начислена", задача выполнена оплата начислена
    4 = "Оплачена", задача выполнена и оплачена исполнителю
    else = "?"
    """
    __tablename__ = 'payment_statuses'

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(16), nullable=False, unique=True)  # Название статуса оплаты

    # Связь с задачами (если будет создана модель Task)
    tasks: Mapped[List["Task"]] = relationship(back_populates="payment_status")

    def __repr__(self) -> str:
        return f"PaymentStatus(id={self.id!r}, name={self.name!r})"


class Task(Base):
    """
    Задачи
    """
    __tablename__ = 'tasks'
    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(128), nullable=False, unique=False)
    #  Имя не уникально, поскольку в разных заказах может быть задача с одним именем например "протестировать"
    description: Mapped[str | None] = mapped_column(String, nullable=True)
    status_id: Mapped[int] = mapped_column(ForeignKey('task_statuses.id'), nullable=True)
    payment_status_id: Mapped[int] = mapped_column(ForeignKey('payment_statuses.id'), nullable=True) #  Пока не отслеживаем
    executor_uuid: Mapped[Optional[UUID]] = mapped_column(ForeignKey('people.uuid'), nullable=True)

    # Запланированное время на выполнение задачи
    planned_duration: Mapped[Optional[timedelta]] = mapped_column(Interval, nullable=True)

    # Фактическое время на выполнение задачи
    actual_duration: Mapped[Optional[timedelta]] = mapped_column(Interval, nullable=True)

    # Дата и время создания задачи
    creation_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    # Дата и время начала выполнения задачи
    start_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    # Планируемая дата и время завершения выполнения задачи
    deadline_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    # Фактическая дата и время завершения выполнения задачи
    end_moment: Mapped[Optional[datetime]] = mapped_column(DateTime, nullable=True)

    # Стоимость выполнения задачи, сколько денег надо заплатить исполнителю, руб
    price: Mapped[int | None] = mapped_column(Integer, nullable=True)

    # Является ли задача шаблонной
    is_template: Mapped[bool] = mapped_column(Boolean, default=False)

    # Добавляем связи с другими таблицами
    status: Mapped["TaskStatus"] = relationship(back_populates="tasks")
    payment_status: Mapped["TaskPaymentStatus"] = relationship(back_populates="tasks")
    timings: Mapped[List["Timing"]] = relationship(back_populates="task")

    # Связь с заказами
    order_serial: Mapped[Optional[str]] = mapped_column(ForeignKey('orders.serial'), nullable=True)
    order: Mapped[Optional["Order"]] = relationship(back_populates="tasks")

    # Ссылка на родительскую задачу
    parent_task_id: Mapped[Optional[int]] = mapped_column(ForeignKey('tasks.id'), nullable=True)

    # Ссылка на корневую задачу
    root_task_id: Mapped[Optional[int]] = mapped_column(ForeignKey('tasks.id'), nullable=True)

    # Связь с исполнителем
    executor: Mapped[Optional["Person"]] = relationship(
        back_populates="tasks",  # Предполагаем, что в Person добавим обратную связь
        foreign_keys="[Task.executor_uuid]"
    )

    # Связи с явным указанием foreign_keys
    parent_task: Mapped["Task"] = relationship(
        "Task",
        remote_side="[Task.id]",
        back_populates="subtasks",
        foreign_keys="[Task.parent_task_id]"
    )
    subtasks: Mapped[List["Task"]] = relationship(
        back_populates="parent_task",
        foreign_keys="[Task.parent_task_id]"  #
    )
    root_task: Mapped["Task"] = relationship(
        "Task",
        remote_side="[Task.id]",
        back_populates="all_tasks_in_hierarchy",
        foreign_keys="[Task.root_task_id]"
    )
    all_tasks_in_hierarchy: Mapped[List["Task"]] = relationship(
        back_populates="root_task",
        foreign_keys="[Task.root_task_id]"  # Используем строковое представление
    )

    def __repr__(self) -> str:
        return f"Task(id={self.id!r}, name={self.name!r})"


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


class Timing(Base):
    """Таблица для хранения времени выполнения задач"""
    __tablename__ = 'timings'

    id: Mapped[int] = mapped_column(primary_key=True, autoincrement=True)
    order_serial: Mapped[str] = mapped_column(ForeignKey('orders.serial'), nullable=False)  # Заказ
    task_id: Mapped[int] = mapped_column(ForeignKey('tasks.id'), nullable=False)  # Задача
    executor_id: Mapped[Optional[int]] = mapped_column(ForeignKey('people.uuid'), nullable=True)  # Исполнитель
    time: Mapped[timedelta] = mapped_column(Interval, nullable=False)  # Потраченное время
    timing_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)  # Дата тайминга

    # Отношения
    order: Mapped["Order"] = relationship(back_populates="timings")
    task: Mapped["Task"] = relationship(back_populates="timings")
    executor: Mapped["Person"] = relationship(
        back_populates="timing_records",
        foreign_keys="[Timing.executor_id]"
    )

    def __repr__(self) -> str:
        return f"Timing(id={self.id!r}, order_serial={self.order_serial!r}, task_id={self.task_id!r})"


class User(AsyncAttrs, Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(primary_key=True)
    username: Mapped[str] = mapped_column(
        String, nullable=False, unique=True, index=True
    )
    hashed_password: Mapped[str] = mapped_column(String, nullable=False)
    email: Mapped[str] = mapped_column(String, nullable=False, unique=True, index=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=lambda: datetime.now(UTC)
    )
    last_login: Mapped[Optional[datetime]] = mapped_column(
        DateTime(timezone=True), nullable=True
    )
    is_admin: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)

    # Связь один-к-одному с Person
    person: Mapped["Person"] = relationship(
        back_populates="user",
        uselist=False  # Обеспечивает 1:1
    )

    def to_dict(self):
        return {
            "id": self.id,
            "username": self.username,
            "email": self.email,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "last_login": self.last_login.isoformat() if self.last_login else None,
            "is_admin": self.is_admin,
        }


'''  
class Equipment_Suppliers(models.Model):  # Поставщик-Оборудование
    equipment = models.OneToOneField(Equipment, on_delete=models.CASCADE)
    supplier = 
    - Supplier_ID(компания)
    - Price_in(наша  входная   цена)
    - Price_out(выходная   цена, розница)
    - Link(ссылка)
'''


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
    
    # Отношения
    sensors: Mapped[List["Sensor"]] = relationship(back_populates="sensor_type")
    
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
    # Один тип отображения может быть у МНОГИХ датчиков (List)
    
    def __str__(self) -> str:
        return self.name


# Промежуточная таблица для связи многие-ко-многим между Sensor и SensorMeasuredValues
sensor_measured_values_association = Table(
    'sensor_measured_values_association',
    Base.metadata,
    Column('sensor_id', Integer, ForeignKey('sensors.id'), primary_key=True),
    Column('measured_value_id', Integer, ForeignKey('sensor_measured_values.id'), primary_key=True)
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
    # Каждый датчик имеет ТОЛЬКО ОДИН тип отображения
    
    # Связь с типом датчика (один-ко-многим)
    sensor_type_id: Mapped[Optional[int]] = mapped_column(ForeignKey('sensor_types.id'), nullable=True)
    # Каждый датчик имеет ТОЛЬКО ОДИН тип
    
    # Отношения
    shape_type: Mapped[Optional["SensorsShapeType"]] = relationship(back_populates="sensors")
    # Один датчик ссылается на один тип отображения
    
    sensor_type: Mapped[Optional["SensorTypes"]] = relationship(back_populates="sensors")
    # Один датчик имеет один тип
    
    # Связь многие-ко-многим с измеряемыми значениями
    measured_values: Mapped[List["SensorMeasuredValues"]] = relationship(
        secondary=sensor_measured_values_association,
        backref="sensors"
    )