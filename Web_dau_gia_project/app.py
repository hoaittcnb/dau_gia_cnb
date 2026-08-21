"""
Web đấu giá từ thiện - Team C&B
Xây dựng theo đặc tả trong file "web_dau_gia project.docx"

Chức năng chính:
  1. Trang giới thiệu
  2. Trang đăng ký
  3. Trang đăng nhập
  4. Trang quên mật khẩu (gửi mật khẩu ngẫu nhiên 8 ký tự qua email)
  5. Màn hình đấu giá (countdown, danh sách sản phẩm, đặt giá, chat realtime)
  6. Trang quyên góp
  7. Trang quản lý của Admin
"""
import os
import io
import csv
import random
import string
import secrets
from datetime import datetime

from flask import (Flask, render_template, request, redirect, url_for, flash,
                   session, jsonify, abort, Response)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, desc
from flask_mail import Mail, Message
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_socketio import SocketIO, emit

# ---------------------------------------------------------------------------
# Cấu hình ứng dụng
# ---------------------------------------------------------------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
ALLOWED_EXT = {'png', 'jpg', 'jpeg', 'gif', 'webp'}

app = Flask(__name__)
app.secret_key = 'doi-thanh-secret-key-cnb-2025'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'instance', 'daugia.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(os.path.join(BASE_DIR, 'instance'), exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

db = SQLAlchemy(app)
mail = Mail(app)
socketio = SocketIO(app, cors_allowed_origins='*')

# last_seen[username] = thời điểm nhận heartbeat gần nhất (dùng để tự động offline)
last_seen = {}
HEARTBEAT_TIMEOUT = 40   # giây - quá thời gian này không có tín hiệu => offline

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    fullname = db.Column(db.String(120), nullable=False)
    password = db.Column(db.String(256), nullable=False)
    is_admin = db.Column(db.Boolean, default=False)
    is_online = db.Column(db.Boolean, default=False)


class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    stt = db.Column(db.Integer, unique=True, nullable=False)   # số thứ tự ngẫu nhiên tăng dần
    name = db.Column(db.String(200), nullable=False)
    description = db.Column(db.String(200))                    # mô tả ngắn (<=30 chữ)
    ext_link = db.Column(db.String(500))                       # link tham khảo trang khác
    image = db.Column(db.String(300))                          # tên file hình
    start_price = db.Column(db.Integer, nullable=False)        # giá khởi điểm
    donor = db.Column(db.String(120), nullable=False)         # người quyên góp
    status = db.Column(db.String(20), default='dong')          # dang | dong | done
    final_price = db.Column(db.Integer)                        # giá chốt đấu giá
    winner = db.Column(db.String(120))                         # người chốt giá
    donate_date = db.Column(db.DateTime, default=datetime.now)


class Bid(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    bidder = db.Column(db.String(120), nullable=False)         # người đấu giá
    amount = db.Column(db.Integer, nullable=False)             # giá đấu
    created_at = db.Column(db.DateTime, default=datetime.now)


class Setting(db.Model):
    """Lưu cấu hình dạng key-value: thời gian đấu giá, mail server, người liên hệ..."""
    key = db.Column(db.String(60), primary_key=True)
    value = db.Column(db.String(500))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def get_setting(key, default=''):
    s = db.session.get(Setting, key)
    return s.value if s and s.value is not None else default


def set_setting(key, value):
    s = db.session.get(Setting, key)
    if s:
        s.value = value
    else:
        db.session.add(Setting(key=key, value=value))


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


def save_image(file_storage):
    """Lưu file hình upload, trả về tên file (hoặc None)."""
    if file_storage and file_storage.filename and allowed_file(file_storage.filename):
        fname = secure_filename(file_storage.filename)
        # tránh trùng tên
        name, ext = os.path.splitext(fname)
        fname = f"{name}_{secrets.token_hex(4)}{ext}"
        file_storage.save(os.path.join(app.config['UPLOAD_FOLDER'], fname))
        return fname
    return None


def docx_to_html(path):
    """Chuyển nội dung file Word (.docx) thành HTML để hiển thị lên trang web.
    Giữ được tiêu đề, in đậm/nghiêng, danh sách, bảng và hình ảnh nhúng trong file."""
    import html as _html
    from docx import Document
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    doc = Document(path)
    img_cache = {}   # rId -> tên file đã lưu trong static/uploads

    def save_embedded(rid):
        if rid in img_cache:
            return img_cache[rid]
        try:
            part = doc.part.related_parts[rid]
        except KeyError:
            return None
        ext = os.path.splitext(part.partname)[1] or '.png'
        fname = f"act_{secrets.token_hex(6)}{ext}"
        with open(os.path.join(app.config['UPLOAD_FOLDER'], fname), 'wb') as f:
            f.write(part.blob)
        img_cache[rid] = fname
        return fname

    def run_html(run):
        # hình ảnh nằm trong run?
        imgs = ''
        for blip in run._element.findall('.//' + qn('a:blip')):
            rid = blip.get(qn('r:embed'))
            fname = save_embedded(rid) if rid else None
            if fname:
                imgs += (f'<img src="/static/uploads/{fname}" '
                         f'alt="hình" style="max-width:100%;border-radius:8px;margin:8px 0">')
        text = _html.escape(run.text or '')
        if text:
            if run.bold:
                text = f'<strong>{text}</strong>'
            if run.italic:
                text = f'<em>{text}</em>'
        return text + imgs

    def para_html(p):
        inner = ''.join(run_html(r) for r in p.runs)
        style = (p.style.name or '').lower()
        if not inner.strip():
            return ''
        if style.startswith('heading 1') or style == 'title':
            return f'<h2>{inner}</h2>'
        if style.startswith('heading 2'):
            return f'<h3>{inner}</h3>'
        if style.startswith('heading'):
            return f'<h4>{inner}</h4>'
        if 'list' in style:
            return f'<li>{inner}</li>'
        return f'<p>{inner}</p>'

    def table_html(tbl):
        rows = ''
        for row in tbl.rows:
            cells = ''.join(f'<td>{ "".join(para_html(p) for p in c.paragraphs) or c.text }</td>'
                            for c in row.cells)
            rows += f'<tr>{cells}</tr>'
        return f'<table class="doc-table">{rows}</table>'

    parts = []
    in_list = False
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            frag = para_html(Paragraph(child, doc))
            is_li = frag.startswith('<li>')
            if is_li and not in_list:
                parts.append('<ul>'); in_list = True
            elif not is_li and in_list:
                parts.append('</ul>'); in_list = False
            if frag:
                parts.append(frag)
        elif child.tag == qn('w:tbl'):
            if in_list:
                parts.append('</ul>'); in_list = False
            parts.append(table_html(Table(child, doc)))
    if in_list:
        parts.append('</ul>')
    return '\n'.join(parts)


def next_stt():
    """Số thứ tự ngẫu nhiên tăng dần, không trùng."""
    max_stt = db.session.query(func.max(Product.stt)).scalar() or 1000
    return max_stt + random.randint(1, 9)


def bid_step(start_price):
    """Bước giá theo giá khởi điểm."""
    if start_price < 100_000:
        return 20_000
    elif start_price < 1_000_000:
        return 30_000
    else:
        return 100_000


def current_price(product):
    """Giá cao nhất hiện tại của sản phẩm (mặc định = giá khởi điểm)."""
    top = db.session.query(func.max(Bid.amount)).filter_by(product_id=product.id).scalar()
    return top if top is not None else product.start_price


def top_bidder(product):
    b = Bid.query.filter_by(product_id=product.id).order_by(desc(Bid.amount)).first()
    return b.bidder if b else None


def configure_mail():
    """Nạp cấu hình mail do admin thiết lập vào flask-mail."""
    app.config['MAIL_SERVER'] = get_setting('mail_server', 'smtp.gmail.com')
    app.config['MAIL_PORT'] = int(get_setting('mail_port', '587') or 587)
    app.config['MAIL_USERNAME'] = get_setting('mail_username', '')
    app.config['MAIL_PASSWORD'] = get_setting('mail_password', '')
    app.config['MAIL_USE_TLS'] = get_setting('mail_use_tls', 'true').lower() == 'true'
    app.config['MAIL_USE_SSL'] = get_setting('mail_use_ssl', 'false').lower() == 'true'
    app.config['MAIL_DEFAULT_SENDER'] = get_setting('mail_username', '')
    mail.init_app(app)


def send_email(to, subject, body):
    """Gửi email, trả về (ok, message)."""
    if not get_setting('mail_username'):
        return False, 'Chưa cấu hình email server trong trang Admin.'
    try:
        configure_mail()
        msg = Message(subject, sender=app.config['MAIL_DEFAULT_SENDER'], recipients=[to])
        msg.body = body
        mail.send(msg)
        return True, 'OK'
    except Exception as e:
        return False, str(e)


def current_user():
    uid = session.get('user_id')
    return db.session.get(User, uid) if uid else None


@app.context_processor
def inject_user():
    """Cho phép mọi template truy cập user đang đăng nhập qua biến cur_user."""
    return {'cur_user': current_user()}


def auction_end():
    v = get_setting('auction_end')
    if not v:
        return None
    try:
        return datetime.fromisoformat(v)
    except ValueError:
        return None


def auction_start():
    v = get_setting('auction_start')
    if not v:
        return None
    try:
        return datetime.fromisoformat(v)
    except ValueError:
        return None


def finalize_auction():
    """Khi hết giờ đấu giá: chốt giá, chuyển status done, gửi email cho người thắng."""
    end = auction_end()
    if not end or datetime.now() < end:
        return
    if get_setting('auction_finalized') == 'true':
        return

    contact_name = get_setting('contact_name', 'Ban tổ chức')
    contact_phone = get_setting('contact_phone', '')

    for p in Product.query.filter_by(status='dang').all():
        top = current_price(p)
        winner = top_bidder(p)
        # chỉ chốt khi có giá đấu lớn hơn giá khởi điểm
        if winner and top > p.start_price:
            p.final_price = top
            p.winner = winner
            p.status = 'done'
            u = User.query.filter_by(username=winner).first()
            if u:
                body = (
                    f"Xin chúc mừng {u.fullname}!\n\n"
                    f"Bạn đã đấu giá thành công vật phẩm: {p.name}\n"
                    f"Giá mua được: {top:,d} VND\n\n"
                    f"Thông tin người liên hệ nhận hàng & chuyển tiền:\n"
                    f"  Họ tên: {contact_name}\n"
                    f"  Điện thoại: {contact_phone}\n\n"
                    f"Đại diện ban tổ chức Cảm ơn bạn đã ủng hộ chương trình!"
                    f"Trân trọng,"
                )
                send_email(u.email, 'Kết quả đấu giá từ thiện', body)
    set_setting('auction_finalized', 'true')
    db.session.commit()


# ---------------------------------------------------------------------------
# Routes - Trang giới thiệu / đăng ký / đăng nhập / quên mật khẩu
# ---------------------------------------------------------------------------
@app.route('/')
def home():
    return render_template('home.html')


@app.route('/activities')
def activities():
    """Trang hoạt động của hội - công khai, không cần đăng nhập.
    Nội dung do admin cập nhật bằng cách upload file Word."""
    content = get_setting('activities_html')
    updated = get_setting('activities_updated')
    return render_template('activities.html', content=content, updated=updated)


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username'].strip()
        email = request.form['email'].strip()
        fullname = request.form['fullname'].strip()
        password = request.form['password']

        if User.query.filter_by(username=username).first():
            flash('Tên đăng nhập đã tồn tại, vui lòng chọn tên khác.', 'error')
        elif User.query.filter_by(email=email).first():
            flash('Email đã tồn tại, vui lòng chọn email khác.', 'error')
        else:
            u = User(username=username, email=email, fullname=fullname,
                     password=generate_password_hash(password))
            db.session.add(u)
            db.session.commit()
            flash('Đăng ký thành công! Vui lòng đăng nhập.', 'success')
            return redirect(url_for('login'))
    return render_template('register.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password']
        u = User.query.filter_by(username=username).first()
        if u and check_password_hash(u.password, password):
            session['user_id'] = u.id
            u.is_online = True
            db.session.commit()
            last_seen[u.username] = datetime.now()
            broadcast_online()      # báo cho mọi người biết có user mới online
            if u.is_admin:
                return redirect(url_for('admin'))
            return redirect(url_for('auction'))
        flash('Tên đăng nhập hoặc mật khẩu không đúng. Vui lòng thử lại.', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    u = current_user()
    if u:
        u.is_online = False
        db.session.commit()
        broadcast_online()      # báo cho mọi người biết user đã thoát
    session.pop('user_id', None)
    flash('Bạn đã thoát khỏi hệ thống.', 'info')
    return redirect(url_for('home'))


@app.route('/forget-password', methods=['GET', 'POST'])
def forget_password():
    if request.method == 'POST':
        email = request.form['email'].strip()
        u = User.query.filter_by(email=email).first()
        if u:
            newpass = ''.join(secrets.choice(string.ascii_letters + string.digits)
                              for _ in range(8))
            u.password = generate_password_hash(newpass)
            db.session.commit()
            ok, msg = send_email(
                u.email, 'Mật khẩu mới - Đấu giá từ thiện',
                f'Xin chào {u.fullname},\n\nMật khẩu mới của bạn là: {newpass}\n\n'
                f'Vui lòng đăng nhập và đổi lại mật khẩu.')
            if ok:
                flash('Mật khẩu mới đã được gửi vào email của bạn. Vui lòng kiểm tra mail.', 'success')
            else:
                flash(f'Không gửi được email ({msg}). Mật khẩu tạm thời của bạn là: {newpass}', 'error')
        else:
            flash('Email không tồn tại trong hệ thống.', 'error')
    return render_template('forget_password.html')


# ---------------------------------------------------------------------------
# Route - Đổi mật khẩu
# ---------------------------------------------------------------------------
@app.route('/change-password', methods=['GET', 'POST'])
def change_password():
    u = current_user()
    if not u:
        return redirect(url_for('login'))
    if request.method == 'POST':
        old = request.form['oldpassword']
        new1 = request.form['newpassword']
        new2 = request.form['confirmpassword']
        if not check_password_hash(u.password, old):
            flash('Mật khẩu cũ không đúng. Vui lòng nhập lại.', 'error')
        elif new1 != new2:
            flash('Hai mật khẩu mới không giống nhau. Vui lòng nhập lại.', 'error')
        else:
            u.password = generate_password_hash(new1)
            db.session.commit()
            flash('Đổi mật khẩu thành công.', 'success')
    return render_template('change_password.html', user=u)


# ---------------------------------------------------------------------------
# Route - Màn hình đấu giá
# ---------------------------------------------------------------------------
@app.route('/auction')
def auction():
    u = current_user()
    if not u:
        flash('Bạn cần đăng nhập trước.', 'info')
        return redirect(url_for('login'))

    finalize_auction()

    products = Product.query.filter_by(status='dang').order_by(Product.stt).all()
    items = []
    for p in products:
        cp = current_price(p)
        items.append({
            'p': p,
            'current': cp,
            'top_bidder': top_bidder(p),
            'step': bid_step(p.start_price),
            'min_next': cp + bid_step(p.start_price),
        })
    end = auction_end()
    return render_template('auction.html', user=u, items=items,
                           auction_end=end.isoformat() if end else '',
                           now=datetime.now())


@app.route('/bid', methods=['POST'])
def bid():
    u = current_user()
    if not u:
        return jsonify({'ok': False, 'msg': 'Chưa đăng nhập'}), 401

    end = auction_end()
    if end and datetime.now() > end:
        return jsonify({'ok': False, 'msg': 'Phiên đấu giá đã kết thúc.'})

    pid = int(request.form['product_id'])
    try:
        amount = int(request.form['amount'])
    except (ValueError, KeyError):
        return jsonify({'ok': False, 'msg': 'Giá không hợp lệ.'})

    p = db.session.get(Product, pid)
    if not p or p.status != 'dang':
        return jsonify({'ok': False, 'msg': 'Sản phẩm không tồn tại hoặc chưa mở đấu giá.'})

    cp = current_price(p)
    step = bid_step(p.start_price)
    min_next = cp + step
    if amount < min_next:
        return jsonify({'ok': False,
                        'msg': f'Vui lòng đặt giá tối thiểu {min_next:,d} VND (bước giá {step:,d}).'})

    db.session.add(Bid(product_id=pid, bidder=u.username, amount=amount))
    db.session.commit()

    # thông báo realtime cho mọi người
    socketio.emit('bid_update', {
        'product_id': pid, 'current': amount, 'bidder': u.username,
        'min_next': amount + step
    })
    return jsonify({'ok': True, 'current': amount, 'bidder': u.username,
                    'min_next': amount + step})


# ---------------------------------------------------------------------------
# Route - Trang quyên góp
# ---------------------------------------------------------------------------
@app.route('/donate', methods=['GET', 'POST'])
def donate():
    u = current_user()
    if not u:
        return redirect(url_for('login'))
    if request.method == 'POST':
        name = request.form['name'].strip()
        description = ' '.join(request.form['description'].split()[:30])  # tối đa 30 chữ
        ext_link = request.form.get('ext_link', '').strip()
        start_price = int(request.form['start_price'])
        image = save_image(request.files.get('image'))
        p = Product(stt=next_stt(), name=name, description=description,
                    ext_link=ext_link, image=image, start_price=start_price,
                    donor=u.username, status='dong')
        db.session.add(p)
        db.session.commit()
        flash('Đã gửi sản phẩm quyên góp thành công! (Không thể chỉnh sửa sau khi gửi)', 'success')
        return redirect(url_for('donate'))

    my_products = Product.query.filter_by(donor=u.username).order_by(Product.stt).all()
    return render_template('donate.html', user=u, my_products=my_products)


# ---------------------------------------------------------------------------
# Route - Trang quản lý Admin
# ---------------------------------------------------------------------------
def require_admin():
    u = current_user()
    if not u:
        return None, redirect(url_for('login'))
    if not u.is_admin:
        abort(403)
    return u, None


@app.route('/admin')
def admin():
    u, resp = require_admin()
    if resp:
        return resp
    finalize_auction()

    products = Product.query.order_by(Product.stt).all()
    prod_view = [{'p': p, 'current': current_price(p), 'top': top_bidder(p)} for p in products]

    bids = Bid.query.order_by(Bid.product_id, desc(Bid.amount)).all()
    bid_view = []
    for b in bids:
        prod = db.session.get(Product, b.product_id)
        bid_view.append({'b': b, 'pname': prod.name if prod else '?', 'stt': prod.stt if prod else ''})

    users = User.query.order_by(User.id).all()

    settings = {
        'auction_start': get_setting('auction_start'),
        'auction_end': get_setting('auction_end'),
        'mail_server': get_setting('mail_server', 'smtp.gmail.com'),
        'mail_port': get_setting('mail_port', '587'),
        'mail_username': get_setting('mail_username'),
        'mail_password': get_setting('mail_password'),
        'mail_use_tls': get_setting('mail_use_tls', 'true'),
        'mail_use_ssl': get_setting('mail_use_ssl', 'false'),
        'contact_name': get_setting('contact_name'),
        'contact_phone': get_setting('contact_phone'),
        'activities_updated': get_setting('activities_updated'),
    }
    return render_template('admin.html', user=u, prod_view=prod_view,
                           bid_view=bid_view, users=users, settings=settings,
                           all_users=users)


@app.route('/admin/product/status', methods=['POST'])
def admin_product_status():
    u, resp = require_admin()
    if resp:
        return resp
    p = db.session.get(Product, int(request.form['product_id']))
    new_status = request.form['status']
    if p and p.status != 'done':          # done thì không đổi được nữa
        if new_status in ('dang', 'dong'):
            p.status = new_status
            # nếu mở lại đấu giá thì cho phép finalize lần sau
            set_setting('auction_finalized', 'false')
            db.session.commit()
            flash('Đã cập nhật trạng thái sản phẩm.', 'success')
    else:
        flash('Sản phẩm đã DONE, không thể đổi trạng thái.', 'error')
    return redirect(url_for('admin') + '#products')


@app.route('/admin/product/add', methods=['POST'])
def admin_product_add():
    u, resp = require_admin()
    if resp:
        return resp
    name = request.form['name'].strip()
    description = ' '.join(request.form['description'].split()[:30])
    ext_link = request.form.get('ext_link', '').strip()
    start_price = int(request.form['start_price'])
    donor = request.form['donor'].strip() or u.username   # admin có thể nhập người quyên góp
    image = save_image(request.files.get('image'))
    p = Product(stt=next_stt(), name=name, description=description, ext_link=ext_link,
                image=image, start_price=start_price, donor=donor, status='dong')
    db.session.add(p)
    db.session.commit()
    flash('Đã thêm sản phẩm.', 'success')
    return redirect(url_for('admin') + '#add-product')


@app.route('/admin/settime', methods=['POST'])
def admin_settime():
    u, resp = require_admin()
    if resp:
        return resp
    start = request.form.get('auction_start', '')
    end = request.form.get('auction_end', '')
    set_setting('auction_start', start)
    set_setting('auction_end', end)
    set_setting('auction_finalized', 'false')
    db.session.commit()
    flash('Đã lưu thời gian đấu giá.', 'success')
    return redirect(url_for('admin') + '#settime')


@app.route('/admin/user/add', methods=['POST'])
def admin_user_add():
    u, resp = require_admin()
    if resp:
        return resp
    username = request.form['username'].strip()
    email = request.form['email'].strip()
    fullname = request.form['fullname'].strip()
    password = request.form['password']
    if User.query.filter_by(username=username).first():
        flash('Tên đăng nhập đã tồn tại.', 'error')
    elif User.query.filter_by(email=email).first():
        flash('Email đã tồn tại.', 'error')
    else:
        db.session.add(User(username=username, email=email, fullname=fullname,
                            password=generate_password_hash(password),
                            is_admin=bool(request.form.get('is_admin'))))
        db.session.commit()
        flash('Đã thêm user mới.', 'success')
    return redirect(url_for('admin') + '#users')


@app.route('/admin/user/delete', methods=['POST'])
def admin_user_delete():
    u, resp = require_admin()
    if resp:
        return resp
    target = db.session.get(User, int(request.form['user_id']))
    if target and target.id != u.id:
        db.session.delete(target)
        db.session.commit()
        flash('Đã xóa user.', 'success')
    else:
        flash('Không thể xóa user này.', 'error')
    return redirect(url_for('admin') + '#users')


@app.route('/admin/mail', methods=['POST'])
def admin_mail():
    u, resp = require_admin()
    if resp:
        return resp
    for key in ('mail_server', 'mail_port', 'mail_username', 'mail_password'):
        set_setting(key, request.form.get(key, ''))
    set_setting('mail_use_tls', 'true' if request.form.get('mail_use_tls') else 'false')
    set_setting('mail_use_ssl', 'true' if request.form.get('mail_use_ssl') else 'false')
    db.session.commit()
    flash('Đã lưu cấu hình email.', 'success')
    return redirect(url_for('admin') + '#mail')


@app.route('/admin/contact', methods=['POST'])
def admin_contact():
    u, resp = require_admin()
    if resp:
        return resp
    set_setting('contact_name', request.form.get('contact_name', ''))
    set_setting('contact_phone', request.form.get('contact_phone', ''))
    db.session.commit()
    flash('Đã lưu thông tin người liên hệ.', 'success')
    return redirect(url_for('admin') + '#contact')


@app.route('/admin/activities', methods=['POST'])
def admin_activities():
    """Admin upload file Word -> nội dung hiển thị lên trang Hoạt động của hội."""
    u, resp = require_admin()
    if resp:
        return resp
    f = request.files.get('docfile')
    if not f or not f.filename.lower().endswith('.docx'):
        flash('Vui lòng chọn file Word định dạng .docx', 'error')
        return redirect(url_for('admin') + '#activities')
    # lưu file gốc rồi chuyển sang HTML
    path = os.path.join(BASE_DIR, 'instance', 'activities.docx')
    f.save(path)
    try:
        html = docx_to_html(path)
    except Exception as e:
        flash(f'Không đọc được file Word: {e}', 'error')
        return redirect(url_for('admin') + '#activities')
    set_setting('activities_html', html)
    set_setting('activities_updated', datetime.now().strftime('%d/%m/%Y %H:%M'))
    db.session.commit()
    flash('Đã cập nhật nội dung trang Hoạt động của hội.', 'success')
    return redirect(url_for('admin') + '#activities')


@app.route('/admin/export-winners')
def admin_export_winners():
    """Xuất danh sách người chiến thắng của ngày đấu giá ra file CSV.
    Thông tin: người đấu giá, vật mua được, giá mua được."""
    u, resp = require_admin()
    if resp:
        return resp
    finalize_auction()

    winners = (Product.query
               .filter(Product.status == 'done', Product.winner.isnot(None))
               .order_by(Product.stt).all())

    buf = io.StringIO()
    buf.write('﻿')          # BOM để Excel đọc đúng tiếng Việt (UTF-8)
    writer = csv.writer(buf)
    writer.writerow(['STT', 'Người đấu giá', 'Vật mua được', 'Giá mua được (VND)'])
    for p in winners:
        wu = User.query.filter_by(username=p.winner).first()
        name = f"{wu.fullname} ({p.winner})" if wu else p.winner
        writer.writerow([p.stt, name, p.name, f"{p.final_price:,d}"])

    fname = f"nguoi_chien_thang_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    return Response(buf.getvalue(), mimetype='text/csv',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})


# ---------------------------------------------------------------------------
# Socket.IO - Chat realtime
# ---------------------------------------------------------------------------
def online_usernames():
    """Danh sách user đang online - dựa theo cột is_online (đăng nhập/đăng xuất),
    không phụ thuộc vào việc mở/đóng từng trang."""
    return [u.username for u in User.query.filter_by(is_online=True).all()]


def broadcast_online():
    """Gửi danh sách online mới nhất cho mọi client đang kết nối."""
    socketio.emit('user_list', online_usernames())


@socketio.on('connect')
def on_connect():
    # Khi bất kỳ trang nào kết nối socket: đánh dấu online + cập nhật heartbeat.
    u = current_user()
    if u:
        last_seen[u.username] = datetime.now()
        if not u.is_online:
            u.is_online = True
            db.session.commit()
    emit('user_list', online_usernames(), broadcast=True)


@socketio.on('heartbeat')
def on_heartbeat():
    """Trang đang mở gửi tín hiệu 'còn sống' định kỳ."""
    u = current_user()
    if u:
        last_seen[u.username] = datetime.now()
        if not u.is_online:      # vừa bị reap nhầm thì cho online lại
            u.is_online = True
            db.session.commit()
            broadcast_online()


def presence_reaper():
    """Tiến trình nền: tự đánh dấu offline các user quá lâu không có heartbeat
    (đóng trình duyệt mà không bấm Thoát)."""
    while True:
        socketio.sleep(20)
        with app.app_context():
            now = datetime.now()
            changed = False
            for u in User.query.filter_by(is_online=True).all():
                ts = last_seen.get(u.username)
                if ts is None or (now - ts).total_seconds() > HEARTBEAT_TIMEOUT:
                    u.is_online = False
                    last_seen.pop(u.username, None)
                    changed = True
            if changed:
                db.session.commit()
                broadcast_online()


@socketio.on('send_message')
def on_message(data):
    u = current_user()
    if u:
        emit('receive_message',
             {'user': u.username, 'message': data.get('message', ''),
              'time': datetime.now().strftime('%H:%M')},
             broadcast=True)


# ---------------------------------------------------------------------------
# Khởi tạo DB + admin mặc định
# ---------------------------------------------------------------------------
def init_db():
    with app.app_context():
        db.create_all()
        # reset trạng thái online (tránh sót cờ online từ lần chạy trước)
        for u in User.query.filter_by(is_online=True).all():
            u.is_online = False
        db.session.commit()
        if not User.query.filter_by(is_admin=True).first():
            db.session.add(User(username='admin', email='admin@cnb.com',
                                fullname='Quản trị viên',
                                password=generate_password_hash('admin123'),
                                is_admin=True))
            db.session.commit()
            print('>> Default admin created: admin / admin123')


if __name__ == '__main__':
    init_db()
    socketio.start_background_task(presence_reaper)   # tự động offline khi đóng trình duyệt
    socketio.run(app, debug=True, host='0.0.0.0', port=8052,
                 allow_unsafe_werkzeug=True)