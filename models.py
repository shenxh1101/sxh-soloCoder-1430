from datetime import datetime, date, timedelta
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False)
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
    monthly_capacity = db.Column(db.Integer, default=1000)
    created_at = db.Column(db.DateTime, default=datetime.now)

    orders = db.relationship('Order', backref='supplier', lazy=True, foreign_keys='Order.supplier_id')

    def get_pending_quantity(self):
        pending = db.session.query(func.sum(Order.quantity)).filter(
            Order.supplier_id == self.id,
            Order.status.in_(['已下单', '已接单', '生产中', '已发货'])
        ).scalar() or 0
        received = db.session.query(func.sum(Receipt.received_quantity)).join(Order).filter(
            Order.supplier_id == self.id,
            Order.status.in_(['已下单', '已接单', '生产中', '已发货', '部分到货', '已到货', '质检中'])
        ).scalar() or 0
        return pending - received

    def get_unsettled_amount(self):
        total = 0
        for o in Order.query.filter(
            Order.supplier_id == self.id,
            Order.status.in_(['质检完成', '已完成'])
        ).all():
            total += o.get_payable_amount()
        paid = db.session.query(func.sum(PaymentRequest.amount)).filter(
            PaymentRequest.supplier_id == self.id,
            PaymentRequest.status == '已付款'
        ).scalar() or 0
        return total - paid

    def get_pending_apply_amount(self):
        total = 0
        for o in Order.query.filter(
            Order.supplier_id == self.id,
            Order.status == '质检完成'
        ).all():
            if not o.payment_requests:
                total += o.get_payable_amount()
        return total

    def get_pending_approve_amount(self):
        return db.session.query(func.sum(PaymentRequest.amount)).filter(
            PaymentRequest.supplier_id == self.id,
            PaymentRequest.status == '待审批'
        ).scalar() or 0

    def get_pending_pay_amount(self):
        return db.session.query(func.sum(PaymentRequest.amount)).filter(
            PaymentRequest.supplier_id == self.id,
            PaymentRequest.status == '已审批'
        ).scalar() or 0

    def get_monthly_performance(self, months=6):
        from collections import OrderedDict
        from dateutil.relativedelta import relativedelta
        result = OrderedDict()
        today = date.today()
        for i in range(months - 1, -1, -1):
            month_date = today - relativedelta(months=i)
            month_start = date(month_date.year, month_date.month, 1)
            next_month_date = month_date + relativedelta(months=1)
            month_end = date(next_month_date.year, next_month_date.month, 1) - timedelta(days=1)

            month_key = month_start.strftime('%Y-%m')
            orders_query = Order.query.filter(
                Order.supplier_id == self.id,
                Order.last_inspected_at >= datetime.combine(month_start, datetime.min.time()),
                Order.last_inspected_at <= datetime.combine(month_end, datetime.max.time())
            )
            completed = [o for o in orders_query.all() if o.status in ['质检完成', '已完成']]

            total_completed = len(completed)
            on_time = sum(1 for o in completed if o.is_on_time())
            first_pass = sum(1 for o in completed if o.is_first_pass() is True)
            inspected = sum(1 for o in completed if o.inspections)
            total_cycle = sum(o.get_processing_days() for o in completed)

            defect_list = []
            for o in completed:
                for insp in o.inspections:
                    if insp.unqualified_quantity > 0 and insp.defect_reasons:
                        defect_list.append({
                            'order_no': o.order_no,
                            'reasons': insp.defect_reasons,
                            'qty': insp.unqualified_quantity,
                            'detail': insp.defect_detail or ''
                        })

            result[month_key] = {
                'completed': total_completed,
                'on_time': on_time,
                'on_time_rate': (on_time / total_completed * 100) if total_completed > 0 else 0,
                'first_pass': first_pass,
                'inspected': inspected,
                'first_pass_rate': (first_pass / inspected * 100) if inspected > 0 else 0,
                'avg_cycle': (total_cycle / total_completed) if total_completed > 0 else 0,
                'total_value': sum(o.get_payable_amount() for o in completed),
                'defects': defect_list
            }
        return result

    def get_monthly_payment_summary(self, months=6):
        from collections import OrderedDict
        from dateutil.relativedelta import relativedelta
        result = OrderedDict()
        today = date.today()
        for i in range(months - 1, -1, -1):
            month_date = today - relativedelta(months=i)
            month_start = date(month_date.year, month_date.month, 1)
            next_month_date = month_date + relativedelta(months=1)
            month_end = date(next_month_date.year, next_month_date.month, 1) - timedelta(days=1)

            month_key = month_start.strftime('%Y-%m')

            orders = Order.query.filter(
                Order.supplier_id == self.id,
                Order.status.in_(['质检完成', '已完成'])
            ).all()

            paid = 0
            unpaid = 0
            overdue_unpaid = 0
            for o in orders:
                if o.last_inspected_at and o.last_inspected_at.date() >= month_start and o.last_inspected_at.date() <= month_end:
                    amt = o.get_payable_amount()
                    pr = o.payment_requests
                    if pr and pr[0].status == '已付款' and pr[0].paid_at and pr[0].paid_at.date() >= month_start and pr[0].paid_at.date() <= month_end:
                        paid += amt
                    else:
                        unpaid += amt
                        if o.agreed_delivery_date + timedelta(days=30) < today:
                            overdue_unpaid += amt

            result[month_key] = {
                'paid': paid,
                'unpaid': unpaid,
                'overdue_unpaid': overdue_unpaid
            }
        return result


class Order(db.Model):
    __tablename__ = 'orders'
    id = db.Column(db.Integer, primary_key=True)
    order_no = db.Column(db.String(40), unique=True, nullable=False)
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    drawing_no = db.Column(db.String(80), nullable=False)
    part_name = db.Column(db.String(120))
    quantity = db.Column(db.Integer, nullable=False)
    unit_price = db.Column(db.Float, nullable=False)
    agreed_delivery_date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(20), default='已下单', nullable=False)
    remark = db.Column(db.String(512))

    created_at = db.Column(db.DateTime, default=datetime.now)
    accepted_at = db.Column(db.DateTime)
    production_at = db.Column(db.DateTime)
    shipped_at = db.Column(db.DateTime)
    first_arrived_at = db.Column(db.DateTime)
    last_arrived_at = db.Column(db.DateTime)
    first_inspected_at = db.Column(db.DateTime)
    last_inspected_at = db.Column(db.DateTime)
    completed_at = db.Column(db.DateTime)

    receipts = db.relationship('Receipt', backref='order', lazy=True, order_by='Receipt.created_at')
    inspections = db.relationship('Inspection', backref='order', lazy=True, order_by='Inspection.created_at')
    payment_requests = db.relationship('PaymentRequest', backref='order', lazy=True)

    def get_total_amount(self):
        return self.quantity * self.unit_price

    def get_total_received(self):
        return db.session.query(func.sum(Receipt.received_quantity)).filter(
            Receipt.order_id == self.id
        ).scalar() or 0

    def get_total_qualified(self):
        return db.session.query(func.sum(Inspection.qualified_quantity)).filter(
            Inspection.order_id == self.id
        ).scalar() or 0

    def get_total_unqualified(self):
        return db.session.query(func.sum(Inspection.unqualified_quantity)).filter(
            Inspection.order_id == self.id
        ).scalar() or 0

    def get_total_inspected(self):
        return self.get_total_qualified() + self.get_total_unqualified()

    def get_payable_amount(self):
        return self.get_total_qualified() * self.unit_price

    def get_remaining_quantity(self):
        return self.quantity - self.get_total_received()

    def get_processing_days(self):
        if self.accepted_at and self.last_inspected_at:
            delta = (self.last_inspected_at.date() - self.accepted_at.date()).days
            return delta if delta >= 0 else 0
        return 0

    def is_on_time(self):
        if self.last_inspected_at:
            return self.last_inspected_at.date() <= self.agreed_delivery_date
        return None

    def is_first_pass(self):
        if not self.inspections:
            return None
        for insp in self.inspections:
            if insp.unqualified_quantity > 0:
                return False
        return True

    def can_create_receipt(self):
        return self.status in ['已发货', '部分到货'] and self.get_remaining_quantity() > 0

    def can_create_inspection(self):
        return self.get_pending_inspection_quantity() > 0

    def get_pending_inspection_quantity(self):
        total_received = self.get_total_received()
        total_inspected = self.get_total_inspected()
        return max(0, total_received - total_inspected)

    def update_status_after_receipt(self):
        total_received = self.get_total_received()
        now = datetime.now()
        if total_received >= self.quantity:
            self.status = '已到货'
            self.last_arrived_at = now
        elif total_received > 0:
            self.status = '部分到货'
            if not self.first_arrived_at:
                self.first_arrived_at = now
            self.last_arrived_at = now

    def update_status_after_inspection(self):
        total_inspected = self.get_total_inspected()
        total_received = self.get_total_received()
        now = datetime.now()
        if total_received >= self.quantity and total_inspected >= self.quantity:
            self.status = '质检完成'
            self.last_inspected_at = now
        elif total_inspected > 0:
            self.status = '质检中'
            if not self.first_inspected_at:
                self.first_inspected_at = now
            self.last_inspected_at = now

    def is_overdue_for_payment(self):
        if self.status == '质检完成' and self.last_inspected_at:
            return (date.today() - self.last_inspected_at.date()).days > 30
        return False

    def get_defect_summary(self):
        summary = []
        for insp in self.inspections:
            if insp.unqualified_quantity > 0:
                parts = [insp.defect_reasons] if insp.defect_reasons else []
                if insp.defect_detail:
                    parts.append(insp.defect_detail)
                summary.append({
                    'inspection_no': insp.inspection_no,
                    'receipt_no': insp.receipt.receipt_no if insp.receipt else '',
                    'unqualified': insp.unqualified_quantity,
                    'reasons': insp.defect_reasons or '',
                    'detail': insp.defect_detail or '',
                    'full_text': '；'.join(p for p in parts if p)
                })
        return summary


class Receipt(db.Model):
    __tablename__ = 'receipts'
    id = db.Column(db.Integer, primary_key=True)
    receipt_no = db.Column(db.String(40), unique=True, nullable=False)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    received_quantity = db.Column(db.Integer, nullable=False)
    waybill_no = db.Column(db.String(80))
    receiver = db.Column(db.String(80))
    remark = db.Column(db.String(512))
    created_at = db.Column(db.DateTime, default=datetime.now)

    inspections = db.relationship('Inspection', backref='receipt', lazy=True)

    def get_pending_inspection(self):
        inspected = db.session.query(func.sum(Inspection.qualified_quantity + Inspection.unqualified_quantity)).filter(
            Inspection.receipt_id == self.id
        ).scalar() or 0
        return self.received_quantity - inspected


class Inspection(db.Model):
    __tablename__ = 'inspections'
    id = db.Column(db.Integer, primary_key=True)
    inspection_no = db.Column(db.String(40), unique=True, nullable=False)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    receipt_id = db.Column(db.Integer, db.ForeignKey('receipts.id'), nullable=False)
    qualified_quantity = db.Column(db.Integer, nullable=False, default=0)
    unqualified_quantity = db.Column(db.Integer, nullable=False, default=0)
    defect_reasons = db.Column(db.String(512))
    defect_detail = db.Column(db.String(512))
    inspector = db.Column(db.String(80))
    remark = db.Column(db.String(512))
    created_at = db.Column(db.DateTime, default=datetime.now)

    def is_first_pass(self):
        return self.unqualified_quantity == 0

    def get_inspection_amount(self):
        return self.qualified_quantity * self.order.unit_price


class PaymentRequest(db.Model):
    __tablename__ = 'payment_requests'
    id = db.Column(db.Integer, primary_key=True)
    request_no = db.Column(db.String(40), unique=True, nullable=False)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    status = db.Column(db.String(20), default='待审批', nullable=False)
    applicant = db.Column(db.String(80))
    approver = db.Column(db.String(80))
    remark = db.Column(db.String(512))

    created_at = db.Column(db.DateTime, default=datetime.now)
    approved_at = db.Column(db.DateTime)
    paid_at = db.Column(db.DateTime)

    supplier = db.relationship('Supplier', backref='payment_requests', lazy=True)
