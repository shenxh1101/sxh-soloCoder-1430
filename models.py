from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # 'enterprise' or 'supplier'
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    supplier = db.relationship('Supplier', backref='users', lazy=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Supplier(db.Model):
    __tablename__ = 'suppliers'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), unique=True, nullable=False)
    contact = db.Column(db.String(80))
    phone = db.Column(db.String(30))
    address = db.Column(db.String(256))
    monthly_capacity = db.Column(db.Integer, default=1000)  # 月产能
    created_at = db.Column(db.DateTime, default=datetime.now)

    orders = db.relationship('Order', backref='supplier', lazy=True, foreign_keys='Order.supplier_id')

    def get_pending_quantity(self):
        from sqlalchemy import func
        result = db.session.query(func.sum(Order.quantity)).filter(
            Order.supplier_id == self.id,
            Order.status.in_(['已下单', '已接单', '生产中', '已发货'])
        ).scalar()
        return result or 0


class Order(db.Model):
    __tablename__ = 'orders'
    id = db.Column(db.Integer, primary_key=True)
    order_no = db.Column(db.String(40), unique=True, nullable=False)
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    drawing_no = db.Column(db.String(80), nullable=False)  # 零件图号
    part_name = db.Column(db.String(120))  # 零件名称
    quantity = db.Column(db.Integer, nullable=False)
    unit_price = db.Column(db.Float, nullable=False)  # 加工单价
    agreed_delivery_date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(20), default='已下单', nullable=False)
    # 状态: 已下单, 已接单, 生产中, 已发货, 已到货, 质检完成, 已完成
    remark = db.Column(db.String(512))

    # 时间戳
    created_at = db.Column(db.DateTime, default=datetime.now)
    accepted_at = db.Column(db.DateTime)
    production_at = db.Column(db.DateTime)
    shipped_at = db.Column(db.DateTime)
    arrived_at = db.Column(db.DateTime)
    inspected_at = db.Column(db.DateTime)
    completed_at = db.Column(db.DateTime)

    # 关联
    inspections = db.relationship('Inspection', backref='order', lazy=True, uselist=False)
    payment_requests = db.relationship('PaymentRequest', backref='order', lazy=True)

    def get_total_amount(self):
        return self.quantity * self.unit_price

    def get_payable_amount(self):
        if self.inspections:
            return self.inspections.qualified_quantity * self.unit_price
        return 0

    def get_processing_days(self):
        if self.accepted_at and self.inspected_at:
            return (self.inspected_at.date() - self.accepted_at.date()).days
        return 0

    def is_on_time(self):
        if self.inspected_at:
            return self.inspected_at.date() <= self.agreed_delivery_date
        return None


class Inspection(db.Model):
    __tablename__ = 'inspections'
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False, unique=True)
    qualified_quantity = db.Column(db.Integer, nullable=False, default=0)
    unqualified_quantity = db.Column(db.Integer, nullable=False, default=0)
    defect_reasons = db.Column(db.String(512))  # 不合格原因，逗号分隔
    inspector = db.Column(db.String(80))
    created_at = db.Column(db.DateTime, default=datetime.now)

    def is_first_pass(self):
        return self.unqualified_quantity == 0


class PaymentRequest(db.Model):
    __tablename__ = 'payment_requests'
    id = db.Column(db.Integer, primary_key=True)
    request_no = db.Column(db.String(40), unique=True, nullable=False)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    status = db.Column(db.String(20), default='待审批', nullable=False)
    # 状态: 待审批, 已审批, 已付款
    applicant = db.Column(db.String(80))
    approver = db.Column(db.String(80))
    remark = db.Column(db.String(512))

    created_at = db.Column(db.DateTime, default=datetime.now)
    approved_at = db.Column(db.DateTime)
    paid_at = db.Column(db.DateTime)

    supplier = db.relationship('Supplier', backref='payment_requests', lazy=True)
