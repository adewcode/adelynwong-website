"""
=============================================================================
ADELYN WONG WEBSITE GENERATOR v4 - WITH PROPERTY DETAIL PAGES
=============================================================================
Reads from:
- Masterbook: Properties, Commercial_Properties, Active Listing sheets
- SEO description: seo.docx in each property folder
- Photos: watermark folder in each property folder
- 3D Tour: 3D_link column in Masterbook
=============================================================================
"""

import pandas as pd
import os
import shutil
from datetime import datetime
import re

# Try to import python-docx for reading Word files
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("⚠️ python-docx not installed. Run: pip install python-docx")

# Try to import Pillow for image compression
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("⚠️ Pillow not installed. Run: pip install Pillow")

# Compression settings
COMPRESS_PHOTOS = True          # Set to False to disable compression
MAX_PHOTO_WIDTH = 1920          # Max width in pixels
MAX_PHOTO_HEIGHT = 1080         # Max height in pixels
JPEG_QUALITY = 80               # Quality 1-100 (80 is good balance)

# =============================================================================
# CONFIGURATION
# =============================================================================

MASTERBOOK_PATH = r"E:\Users\ade\OneDrive\MyRealEstateBiz\RealEstate_Database\masterbook_database.xlsm"
PHOTO_BASE_PATH = r"E:\Users\ade\OneDrive\MyRealEstateBiz"
OUTPUT_FOLDER = r"E:\phython_automation_github"
MAX_PHOTOS_PER_PROPERTY = 10
MAX_PHOTOS_DETAIL_PAGE = 15  # Reduced to save storage

# Contact details
WHATSAPP = "60176846282"
WECHAT_ID = "adelynwong80"
EMAIL = "adelynwonglive@gmail.com"

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def read_seo_docx(property_folder):
    """Read SEO description from seo.docx in property folder"""
    if not property_folder or pd.isna(property_folder):
        return ""
    
    if not HAS_DOCX:
        return ""
    
    seo_path = os.path.join(PHOTO_BASE_PATH, str(property_folder), 'seo.docx')
    
    if not os.path.exists(seo_path):
        return ""
    
    try:
        doc = Document(seo_path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        return '\n\n'.join(paragraphs)
    except Exception as e:
        print(f"   ⚠️ Could not read seo.docx: {e}")
        return ""


def get_shared_styles():
    return '''
        :root {
            --primary-dark: #1a1a1a;
            --primary-gold: #C9A55C;
            --text-dark: #2d2d2d;
            --text-light: #666666;
            --bg-light: #fafafa;
            --bg-cream: #fdf8f3;
            --bg-white: #ffffff;
            --green-whatsapp: #25D366;
            --shadow-soft: 0 4px 20px rgba(0,0,0,0.08);
            --shadow-hover: 0 8px 30px rgba(0,0,0,0.12);
            --font-display: 'Cormorant Garamond', Georgia, serif;
            --font-body: 'Montserrat', -apple-system, sans-serif;
            --transition: all 0.3s ease;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: var(--font-body); color: var(--text-dark); line-height: 1.6; background: var(--bg-white); }

        .navbar { position: fixed; top: 0; left: 0; right: 0; z-index: 1000; background: rgba(255,255,255,0.98); backdrop-filter: blur(10px); box-shadow: 0 1px 0 rgba(0,0,0,0.05); }
        .nav-container { max-width: 1400px; margin: 0 auto; padding: 0 2rem; display: flex; justify-content: space-between; align-items: center; height: 70px; }
        .logo { font-family: var(--font-display); font-size: 1.5rem; font-weight: 600; color: var(--primary-dark); text-decoration: none; letter-spacing: 2px; }
        .logo span { color: var(--primary-gold); }
        .nav-links { display: flex; gap: 2.5rem; list-style: none; }
        .nav-links a { font-size: 0.85rem; font-weight: 500; color: var(--text-dark); text-decoration: none; letter-spacing: 1px; text-transform: uppercase; transition: var(--transition); }
        .nav-links a:hover { color: var(--primary-gold); }
        .mobile-toggle { display: none; background: none; border: none; font-size: 1.5rem; cursor: pointer; }

        .footer { background: var(--primary-dark); color: rgba(255,255,255,0.7); padding: 3rem 2rem; text-align: center; }
        .footer-logo { font-family: var(--font-display); font-size: 1.5rem; color: white; margin-bottom: 1rem; }
        .footer-logo span { color: var(--primary-gold); }
        .footer-text { font-size: 0.85rem; margin-bottom: 0.5rem; }
        .footer-ren { font-size: 0.75rem; color: var(--primary-gold); }

        .whatsapp-float { position: fixed; bottom: 20px; right: 20px; width: 60px; height: 60px; background: var(--green-whatsapp); border-radius: 50%; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 20px rgba(37, 211, 102, 0.4); z-index: 999; transition: var(--transition); text-decoration: none; }
        .whatsapp-float:hover { transform: scale(1.1); }
        .whatsapp-float svg { width: 32px; height: 32px; fill: white; }

        @media (max-width: 768px) {
            .nav-links { display: none; }
            .mobile-toggle { display: block; }
        }
    '''


def get_location_filter(location):
    loc = str(location).lower()
    if 'eco park' in loc or 'ecopark' in loc:
        return 'setia-eco-park'
    elif 'eco ardence' in loc or 'ardence' in loc:
        return 'eco-ardence'
    elif 'setia city' in loc or 'setia alam' in loc or 'edusentral' in loc or 'trefoil' in loc:
        return 'setia-city'
    else:
        return 'other'


def get_property_icon(property_type):
    ptype = str(property_type).lower()
    if 'apartment' in ptype or 'condo' in ptype:
        return '🏢'
    elif 'semi' in ptype or 'bungalow' in ptype or 'house' in ptype:
        return '🏡'
    elif 'townhouse' in ptype or 'terrace' in ptype:
        return '🏘️'
    elif 'shop' in ptype or 'office' in ptype or 'commercial' in ptype:
        return '🏪'
    else:
        return '🏠'


def format_price(price, listing_type):
    if pd.isna(price) or price == 0:
        return "Contact for price"
    price = int(price)
    if listing_type == 'Sale':
        if price >= 1000000:
            return f"RM {price/1000000:.2f}M"
        else:
            return f"RM {price:,}"
    else:
        return f"RM {price:,}/month"


def compress_and_copy_photo(src_path, dst_path):
    """Compress photo and save to destination"""
    if not HAS_PIL or not COMPRESS_PHOTOS:
        # Just copy if Pillow not available or compression disabled
        shutil.copy2(src_path, dst_path)
        return
    
    try:
        with Image.open(src_path) as img:
            # Convert to RGB if necessary (for PNG with transparency)
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            
            # Resize if larger than max dimensions
            original_size = img.size
            img.thumbnail((MAX_PHOTO_WIDTH, MAX_PHOTO_HEIGHT), Image.LANCZOS)
            
            # Save as JPEG with compression
            dst_path_jpg = dst_path.rsplit('.', 1)[0] + '.jpg'
            img.save(dst_path_jpg, 'JPEG', quality=JPEG_QUALITY, optimize=True)
            
            return dst_path_jpg
    except Exception as e:
        # If compression fails, just copy original
        shutil.copy2(src_path, dst_path)
        return dst_path


def copy_property_photos(property_folder, prop_id, output_folder, max_photos=10):
    photos_copied = []
    
    if not property_folder or pd.isna(property_folder):
        return photos_copied
    
    source_folder = os.path.join(PHOTO_BASE_PATH, str(property_folder), 'watermark')
    
    if not os.path.exists(source_folder):
        return photos_copied
    
    dest_folder = os.path.join(output_folder, 'photos', str(prop_id))
    os.makedirs(dest_folder, exist_ok=True)
    
    image_extensions = ('.jpg', '.jpeg', '.png', '.webp')
    photo_files = [f for f in os.listdir(source_folder) if f.lower().endswith(image_extensions)]
    photo_files.sort()
    photo_files = photo_files[:max_photos]
    
    for photo in photo_files:
        src = os.path.join(source_folder, photo)
        # Always save as .jpg for consistency
        photo_name = photo.rsplit('.', 1)[0] + '.jpg'
        dst = os.path.join(dest_folder, photo_name)
        try:
            compress_and_copy_photo(src, dst)
            photos_copied.append(photo_name)
        except:
            pass
    
    return photos_copied


# =============================================================================
# MAIN PAGE TEMPLATE
# =============================================================================

def get_main_template():
    return '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Adelyn Wong | Real Estate Matchmaker Since 2012</title>
    <meta name="description" content="Adelyn Wong - Real Estate Matchmaker specializing in Setia Alam, Setia Eco Park, Eco Ardence & Klang. 12+ years experience.">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600;700&family=Montserrat:wght@300;400;500;600&display=swap" rel="stylesheet">
    <style>
        ''' + get_shared_styles() + '''

        .stats { padding: 3rem 2rem; background: var(--bg-white); margin-top: 70px; }
        .stats-container { max-width: 1200px; margin: 0 auto; display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.5rem; }
        .stat-item { text-align: center; padding: 1.5rem; background: var(--bg-light); border-radius: 8px; border-left: 3px solid var(--primary-gold); }
        .stat-number { font-family: var(--font-display); font-size: 2.5rem; font-weight: 600; color: var(--primary-dark); }
        .stat-number span { color: var(--primary-gold); }
        .stat-label { font-size: 0.75rem; letter-spacing: 1px; text-transform: uppercase; color: var(--text-light); margin-top: 0.25rem; }

        .listings { padding: 4rem 2rem; background: var(--bg-white); }
        .section-header { text-align: center; margin-bottom: 2.5rem; }
        .section-title { font-family: var(--font-display); font-size: 2.2rem; font-weight: 400; color: var(--primary-dark); margin-bottom: 0.5rem; }
        .section-subtitle { font-size: 0.9rem; color: var(--text-light); }

        .filter-tabs { display: flex; justify-content: center; flex-wrap: wrap; gap: 0.5rem; margin-bottom: 2.5rem; }
        .filter-tab { padding: 0.6rem 1.2rem; border: 1px solid #e0e0e0; border-radius: 25px; background: var(--bg-white); font-family: var(--font-body); font-size: 0.8rem; cursor: pointer; transition: var(--transition); }
        .filter-tab:hover, .filter-tab.active { background: var(--primary-dark); color: var(--bg-white); border-color: var(--primary-dark); }

        .property-grid { max-width: 1400px; margin: 0 auto; display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 1.5rem; }
        .property-card { background: var(--bg-white); border-radius: 8px; overflow: hidden; box-shadow: var(--shadow-soft); transition: var(--transition); cursor: pointer; text-decoration: none; color: inherit; display: block; }
        .property-card:hover { transform: translateY(-3px); box-shadow: var(--shadow-hover); }
        
        .property-image { position: relative; height: 200px; background: linear-gradient(135deg, #e8e8e8 0%, #f5f5f5 100%); overflow: hidden; }
        .carousel { position: relative; width: 100%; height: 100%; }
        .carousel-inner { display: flex; transition: transform 0.3s ease; height: 100%; }
        .carousel-item { min-width: 100%; height: 100%; }
        .carousel-item img { width: 100%; height: 100%; object-fit: cover; }
        .carousel-btn { position: absolute; top: 50%; transform: translateY(-50%); background: rgba(255,255,255,0.9); border: none; width: 32px; height: 32px; border-radius: 50%; cursor: pointer; font-size: 1rem; display: flex; align-items: center; justify-content: center; z-index: 10; transition: var(--transition); }
        .carousel-btn:hover { background: var(--primary-gold); color: white; }
        .carousel-btn.prev { left: 8px; }
        .carousel-btn.next { right: 8px; }
        .carousel-dots { position: absolute; bottom: 8px; left: 50%; transform: translateX(-50%); display: flex; gap: 4px; }
        .carousel-dot { width: 6px; height: 6px; border-radius: 50%; background: rgba(255,255,255,0.5); cursor: pointer; }
        .carousel-dot.active { background: white; }
        .photo-count { position: absolute; top: 8px; right: 8px; background: rgba(0,0,0,0.6); color: white; padding: 2px 8px; border-radius: 4px; font-size: 0.7rem; }
        
        .property-image .icon { font-size: 3.5rem; opacity: 0.5; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); }
        
        .property-badges { position: absolute; top: 0.75rem; left: 0.75rem; display: flex; gap: 0.5rem; z-index: 5; }
        .badge { padding: 0.3rem 0.6rem; border-radius: 4px; font-size: 0.65rem; font-weight: 600; letter-spacing: 0.5px; text-transform: uppercase; }
        .badge-status { background: #22c55e; color: white; }
        .badge-status.let-out { background: #94a3b8; }
        .badge-type { background: var(--primary-dark); color: white; }
        .badge-type.for-sale { background: var(--primary-gold); color: white; }
        
        .property-content { padding: 1.25rem; }
        .property-location { font-size: 0.7rem; letter-spacing: 1px; text-transform: uppercase; color: var(--primary-gold); margin-bottom: 0.4rem; }
        .property-type { font-family: var(--font-display); font-size: 1.1rem; font-weight: 500; color: var(--primary-dark); margin-bottom: 0.75rem; }
        .property-specs { display: flex; gap: 0.75rem; margin-bottom: 0.75rem; font-size: 0.8rem; color: var(--text-light); }
        .property-price { font-family: var(--font-display); font-size: 1.3rem; font-weight: 600; color: var(--primary-dark); margin-bottom: 1rem; }

        .property-actions { display: flex; gap: 0.5rem; }
        .btn-view { flex: 1; display: flex; align-items: center; justify-content: center; gap: 0.4rem; padding: 0.75rem; background: var(--primary-dark); color: white; border: none; border-radius: 6px; font-size: 0.8rem; font-weight: 500; text-decoration: none; transition: var(--transition); }
        .btn-view:hover { background: var(--primary-gold); }
        .btn-closed { flex: 1; display: flex; align-items: center; justify-content: center; padding: 0.75rem; background: #94a3b8; color: white; border: none; border-radius: 6px; font-size: 0.8rem; font-weight: 500; }

        @media (max-width: 768px) {
            .stats-container { grid-template-columns: repeat(2, 1fr); }
            .property-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <nav class="navbar">
        <div class="nav-container">
            <a href="index.html" class="logo">ADELYN<span>WONG</span></a>
            <ul class="nav-links">
                <li><a href="index.html">Home</a></li>
                <li><a href="#listings">Listings</a></li>
                <li><a href="#contact">Contact</a></li>
            </ul>
            <button class="mobile-toggle">☰</button>
        </div>
    </nav>

    <section class="stats">
        <div class="stats-container">
            <div class="stat-item">
                <div class="stat-number">12<span>+</span></div>
                <div class="stat-label">Years Experience</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">{total_properties}<span>+</span></div>
                <div class="stat-label">Properties Managed</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">{active_listings}</div>
                <div class="stat-label">Active Listings</div>
            </div>
            <div class="stat-item">
                <div class="stat-number">{closed_deals}<span>+</span></div>
                <div class="stat-label">Deals Closed</div>
            </div>
        </div>
    </section>

    <section class="listings" id="listings">
        <div class="section-header">
            <h2 class="section-title">Property Listings</h2>
            <p class="section-subtitle">Find your perfect home in Setia Alam, Eco Park & Eco Ardence</p>
        </div>
        
        <div class="filter-tabs">
            <button class="filter-tab active" data-filter="all">All</button>
            <button class="filter-tab" data-filter="rent">For Rent</button>
            <button class="filter-tab" data-filter="sale">For Sale</button>
            <button class="filter-tab" data-filter="setia-eco-park">Setia Eco Park</button>
            <button class="filter-tab" data-filter="eco-ardence">Eco Ardence</button>
            <button class="filter-tab" data-filter="setia-city">Setia City</button>
        </div>
        
        <div class="property-grid">
{property_cards}
        </div>
    </section>

    <section class="contact" id="contact" style="padding: 4rem 2rem; background: var(--bg-cream);">
        <div style="max-width: 800px; margin: 0 auto; text-align: center;">
            <h2 class="section-title">Get In Touch</h2>
            <p class="section-subtitle">Ready to find your dream home? Contact me today!</p>
            <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 1rem; margin-top: 2rem;">
                <a href="https://wa.me/{whatsapp}" style="display: inline-flex; align-items: center; gap: 0.5rem; padding: 1rem 2rem; border-radius: 50px; font-weight: 500; text-decoration: none; background: var(--green-whatsapp); color: white;" target="_blank">WhatsApp</a>
                <a href="weixin://dl/chat?{wechat_id}" style="display: inline-flex; align-items: center; gap: 0.5rem; padding: 1rem 2rem; border-radius: 50px; font-weight: 500; text-decoration: none; background: #07c160; color: white;">WeChat: {wechat_id}</a>
            </div>
        </div>
    </section>

    <footer class="footer">
        <div class="footer-logo">ADELYN<span>WONG</span></div>
        <p class="footer-text">Connecting People to Real Estate Since 2012</p>
        <p class="footer-ren">REN 03144 | CID Realtors</p>
        <p class="footer-text" style="margin-top: 1rem; font-size: 0.75rem;">Updated: {generated_date}</p>
    </footer>

    <a href="https://wa.me/{whatsapp}" class="whatsapp-float" target="_blank">
        <svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
    </a>

    <script>
        document.querySelectorAll('.filter-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                const filter = tab.dataset.filter;
                document.querySelectorAll('.property-card').forEach(card => {
                    if (filter === 'all') {
                        card.style.display = 'block';
                    } else {
                        const matches = card.dataset.type === filter || card.dataset.location === filter;
                        card.style.display = matches ? 'block' : 'none';
                    }
                });
            });
        });

        document.querySelectorAll('.carousel').forEach(carousel => {
            const inner = carousel.querySelector('.carousel-inner');
            const items = carousel.querySelectorAll('.carousel-item');
            const dots = carousel.querySelectorAll('.carousel-dot');
            const prevBtn = carousel.querySelector('.carousel-btn.prev');
            const nextBtn = carousel.querySelector('.carousel-btn.next');
            let current = 0;
            const total = items.length;

            function goTo(index) {
                if (index < 0) index = total - 1;
                if (index >= total) index = 0;
                current = index;
                inner.style.transform = 'translateX(-' + (current * 100) + '%)';
                dots.forEach((dot, i) => dot.classList.toggle('active', i === current));
            }

            if (prevBtn) prevBtn.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); goTo(current - 1); });
            if (nextBtn) nextBtn.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); goTo(current + 1); });
            dots.forEach((dot, i) => dot.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); goTo(i); }));
        });
    </script>
</body>
</html>'''


# =============================================================================
# DETAIL PAGE TEMPLATE
# =============================================================================

def get_detail_template():
    return '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{property_title} | Adelyn Wong</title>
    <meta name="description" content="{meta_description}">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600;700&family=Montserrat:wght@300;400;500;600&display=swap" rel="stylesheet">
    <style>
        ''' + get_shared_styles() + '''

        .detail-container { max-width: 1200px; margin: 0 auto; padding: 90px 2rem 4rem; }
        
        .back-btn { display: inline-flex; align-items: center; gap: 0.5rem; color: var(--text-light); text-decoration: none; font-size: 0.9rem; margin-bottom: 1.5rem; transition: var(--transition); }
        .back-btn:hover { color: var(--primary-gold); }
        
        .hero-gallery { display: grid; grid-template-columns: 2fr 1fr; gap: 0.5rem; margin-bottom: 2rem; border-radius: 12px; overflow: hidden; max-height: 500px; }
        .hero-main { position: relative; }
        .hero-main img { width: 100%; height: 100%; object-fit: cover; cursor: pointer; }
        .hero-side { display: grid; grid-template-rows: 1fr 1fr; gap: 0.5rem; }
        .hero-side img { width: 100%; height: 100%; object-fit: cover; cursor: pointer; }
        .hero-more { position: relative; }
        .hero-more::after { content: '+{more_photos} photos'; position: absolute; inset: 0; background: rgba(0,0,0,0.5); color: white; display: flex; align-items: center; justify-content: center; font-weight: 600; cursor: pointer; }
        
        .property-header { display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 1rem; margin-bottom: 2rem; }
        .property-info h1 { font-family: var(--font-display); font-size: 2rem; font-weight: 600; color: var(--primary-dark); margin-bottom: 0.5rem; }
        .property-info .location { color: var(--primary-gold); font-size: 1rem; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 1rem; }
        .property-info .specs { display: flex; gap: 1.5rem; font-size: 1rem; color: var(--text-light); }
        
        .property-price-box { text-align: right; }
        .property-price-box .price { font-family: var(--font-display); font-size: 2.5rem; font-weight: 700; color: var(--primary-dark); }
        .property-price-box .price-label { font-size: 0.85rem; color: var(--text-light); }
        .property-price-box .badge { display: inline-block; padding: 0.4rem 1rem; border-radius: 20px; font-size: 0.75rem; font-weight: 600; text-transform: uppercase; margin-top: 0.5rem; }
        .badge-rent { background: var(--primary-dark); color: white; }
        .badge-sale { background: var(--primary-gold); color: white; }
        
        .action-buttons { display: flex; gap: 1rem; margin: 2rem 0; flex-wrap: wrap; }
        .action-btn { display: inline-flex; align-items: center; justify-content: center; gap: 0.5rem; padding: 1rem 2rem; border-radius: 8px; font-weight: 600; text-decoration: none; transition: var(--transition); border: none; cursor: pointer; font-size: 1rem; }
        .btn-whatsapp { background: var(--green-whatsapp); color: white; flex: 1; min-width: 200px; }
        .btn-whatsapp:hover { background: #1da851; transform: translateY(-2px); }
        .btn-tour { background: var(--primary-dark); color: white; }
        .btn-tour:hover { background: var(--primary-gold); }
        .btn-video { background: #ff0000; color: white; }
        .btn-video:hover { background: #cc0000; }
        
        .content-grid { display: grid; grid-template-columns: 2fr 1fr; gap: 3rem; }
        
        .description-section { background: var(--bg-light); border-radius: 12px; padding: 1.5rem; }
        .description-section h2 { font-family: var(--font-display); font-size: 1.5rem; margin-bottom: 1rem; color: var(--primary-dark); }
        .description-section p { color: var(--text-light); line-height: 1.8; margin-bottom: 1rem; white-space: pre-line; }
        
        .details-box { background: var(--bg-light); border-radius: 12px; padding: 1.5rem; height: fit-content; }
        .details-box h3 { font-family: var(--font-display); font-size: 1.2rem; margin-bottom: 1rem; color: var(--primary-dark); }
        .detail-row { display: flex; justify-content: space-between; padding: 0.75rem 0; border-bottom: 1px solid #e0e0e0; }
        .detail-row:last-child { border-bottom: none; }
        .detail-label { color: var(--text-light); }
        .detail-value { font-weight: 500; color: var(--primary-dark); }
        
        .tour-section { margin-top: 3rem; }
        .tour-section h2 { font-family: var(--font-display); font-size: 1.5rem; margin-bottom: 1.5rem; color: var(--primary-dark); }
        .tour-embed { width: 100%; height: 500px; border-radius: 12px; overflow: hidden; background: var(--bg-light); }
        .tour-embed iframe { width: 100%; height: 100%; border: none; }
        
        .gallery-section { margin-top: 3rem; }
        .gallery-section h2 { font-family: var(--font-display); font-size: 1.5rem; margin-bottom: 1.5rem; color: var(--primary-dark); }
        .gallery-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 1rem; }
        .gallery-grid img { width: 100%; height: 200px; object-fit: cover; border-radius: 8px; cursor: pointer; transition: var(--transition); }
        .gallery-grid img:hover { transform: scale(1.02); }
        
        .lightbox { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.95); z-index: 2000; align-items: center; justify-content: center; }
        .lightbox.active { display: flex; }
        .lightbox img { max-width: 90%; max-height: 90%; object-fit: contain; }
        .lightbox-close { position: absolute; top: 20px; right: 30px; color: white; font-size: 2rem; cursor: pointer; }
        .lightbox-nav { position: absolute; top: 50%; transform: translateY(-50%); color: white; font-size: 3rem; cursor: pointer; padding: 1rem; }
        .lightbox-prev { left: 20px; }
        .lightbox-next { right: 20px; }
        
        @media (max-width: 768px) {
            .hero-gallery { grid-template-columns: 1fr; max-height: none; }
            .hero-side { grid-template-columns: 1fr 1fr; }
            .property-header { flex-direction: column; }
            .property-price-box { text-align: left; }
            .content-grid { grid-template-columns: 1fr; }
            .action-buttons { flex-direction: column; }
            .action-btn { width: 100%; }
        }
    </style>
</head>
<body>
    <nav class="navbar">
        <div class="nav-container">
            <a href="index.html" class="logo">ADELYN<span>WONG</span></a>
            <ul class="nav-links">
                <li><a href="index.html">Home</a></li>
                <li><a href="index.html#listings">Listings</a></li>
                <li><a href="index.html#contact">Contact</a></li>
            </ul>
            <button class="mobile-toggle">☰</button>
        </div>
    </nav>

    <div class="detail-container">
        <a href="index.html" class="back-btn">← Back to Listings</a>
        
        <div class="hero-gallery">
            <div class="hero-main">
                <img src="{hero_image}" alt="{property_title}" onclick="openLightbox(0)">
            </div>
            <div class="hero-side">
                <img src="{side_image_1}" alt="{property_title}" onclick="openLightbox(1)">
                <div class="hero-more" onclick="openLightbox(2)">
                    <img src="{side_image_2}" alt="{property_title}">
                </div>
            </div>
        </div>
        
        <div class="property-header">
            <div class="property-info">
                <h1>{property_type}</h1>
                <p class="location">{location}</p>
                <div class="specs">
                    <span>🛏️ {beds}</span>
                    <span>🚿 {baths}</span>
                    <span>📐 {sqft} sqft</span>
                </div>
            </div>
            <div class="property-price-box">
                <div class="price">{price}</div>
                <div class="price-label">{price_label}</div>
                <span class="badge {badge_class}">{listing_type}</span>
            </div>
        </div>
        
        <div class="action-buttons">
            <a href="{whatsapp_link}" class="action-btn btn-whatsapp" target="_blank">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
                Enquire on WhatsApp
            </a>
            {tour_button}
            {video_button}
        </div>
        
        <div class="content-grid">
            <div class="description-section">
                <h2>About This Property</h2>
                <p>{description}</p>
            </div>
            
            <div class="details-box">
                <h3>Property Details</h3>
                <div class="detail-row">
                    <span class="detail-label">Property Type</span>
                    <span class="detail-value">{property_type}</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Bedrooms</span>
                    <span class="detail-value">{beds}</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Bathrooms</span>
                    <span class="detail-value">{baths}</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Built-up Size</span>
                    <span class="detail-value">{sqft} sqft</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Furnishing</span>
                    <span class="detail-value">{furnishing}</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Tenure</span>
                    <span class="detail-value">{tenure}</span>
                </div>
            </div>
        </div>
        
        {tour_section}
        
        <div class="gallery-section">
            <h2>Photo Gallery</h2>
            <div class="gallery-grid">
                {gallery_images}
            </div>
        </div>
    </div>

    <footer class="footer">
        <div class="footer-logo">ADELYN<span>WONG</span></div>
        <p class="footer-text">Connecting People to Real Estate Since 2012</p>
        <p class="footer-ren">REN 03144 | CID Realtors</p>
    </footer>

    <a href="{whatsapp_link}" class="whatsapp-float" target="_blank">
        <svg viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
    </a>

    <div class="lightbox" id="lightbox">
        <span class="lightbox-close" onclick="closeLightbox()">×</span>
        <span class="lightbox-nav lightbox-prev" onclick="changeSlide(-1)">‹</span>
        <img src="" id="lightbox-img">
        <span class="lightbox-nav lightbox-next" onclick="changeSlide(1)">›</span>
    </div>

    <script>
        const photos = [{photos_json}];
        let currentSlide = 0;

        function openLightbox(index) {
            currentSlide = index;
            document.getElementById('lightbox-img').src = photos[currentSlide];
            document.getElementById('lightbox').classList.add('active');
            document.body.style.overflow = 'hidden';
        }

        function closeLightbox() {
            document.getElementById('lightbox').classList.remove('active');
            document.body.style.overflow = '';
        }

        function changeSlide(dir) {
            currentSlide += dir;
            if (currentSlide < 0) currentSlide = photos.length - 1;
            if (currentSlide >= photos.length) currentSlide = 0;
            document.getElementById('lightbox-img').src = photos[currentSlide];
        }

        document.addEventListener('keydown', (e) => {
            if (!document.getElementById('lightbox').classList.contains('active')) return;
            if (e.key === 'Escape') closeLightbox();
            if (e.key === 'ArrowLeft') changeSlide(-1);
            if (e.key === 'ArrowRight') changeSlide(1);
        });
    </script>
</body>
</html>'''


# =============================================================================
# GENERATE PROPERTY CARD
# =============================================================================

def generate_property_card(row, output_folder):
    prop_id = row.get('Property_ID', '')
    location = row.get('Location', 'Unknown Location')
    property_type = row.get('Property type', 'Property')
    listing_type = row.get('Listing_Type', 'Rent')
    beds = row.get('Bedrooms', 0)
    baths = row.get('Bathrooms', 0)
    sqft = row.get('Built-up', 0)
    rent_price = row.get('Rent_RM', 0) if not pd.isna(row.get('Rent_RM', 0)) else row.get('Rental price', 0)
    sale_price = row.get('Sale_RM', 0) if not pd.isna(row.get('Sale_RM', 0)) else row.get('Selling price', 0)
    ads_status = row.get('Ads_Status', 'In Listing')
    property_folder = row.get('property_folder', '')
    
    photos = copy_property_photos(property_folder, prop_id, output_folder, MAX_PHOTOS_PER_PROPERTY)
    
    if listing_type == 'Sale' and not pd.isna(sale_price) and sale_price > 0:
        price_html = format_price(sale_price, 'Sale')
        type_badge = "For Sale"
        type_class = "for-sale"
    else:
        price_html = format_price(rent_price, 'Rent')
        type_badge = "For Rent"
        type_class = ""
    
    if pd.isna(beds):
        beds_str = '0'
    elif str(beds).lower() == 'studio':
        beds_str = 'Studio'
    else:
        try:
            beds_str = str(int(beds))
        except:
            beds_str = str(beds)
    
    baths_str = str(int(baths)) if not pd.isna(baths) and str(baths).replace('.','').isdigit() else '0'
    sqft_str = f"{int(sqft):,}" if not pd.isna(sqft) and str(sqft).replace('.','').replace(',','').isdigit() and float(sqft) > 0 else "N/A"
    
    location_filter = get_location_filter(location)
    type_filter = 'sale' if listing_type == 'Sale' else 'rent'
    
    ads_lower = str(ads_status).lower()
    if 'let out' in ads_lower:
        status_badge = '<span class="badge badge-status let-out">LET OUT</span>'
        is_closed = True
    elif 'sold' in ads_lower:
        status_badge = '<span class="badge badge-status let-out">SOLD</span>'
        is_closed = True
    else:
        status_badge = '<span class="badge badge-status">AVAILABLE</span>'
        is_closed = False
    
    if photos:
        photo_items = ""
        photo_dots = ""
        for i, photo in enumerate(photos):
            photo_items += f'<div class="carousel-item"><img src="photos/{prop_id}/{photo}" alt="{location}" loading="lazy"></div>'
            active_class = "active" if i == 0 else ""
            photo_dots += f'<span class="carousel-dot {active_class}"></span>'
        
        image_html = f'''
                <div class="carousel">
                    <div class="carousel-inner">{photo_items}</div>
                    <button class="carousel-btn prev">‹</button>
                    <button class="carousel-btn next">›</button>
                    <div class="carousel-dots">{photo_dots}</div>
                    <span class="photo-count">{len(photos)} photos</span>
                </div>'''
    else:
        icon = get_property_icon(property_type)
        image_html = f'<span class="icon">{icon}</span>'
    
    if beds_str == 'Studio':
        beds_display = "🛏️ Studio"
    else:
        beds_display = f"🛏️ {beds_str} Beds"
    
    if is_closed:
        card_link_start = f'<div class="property-card" data-type="{type_filter}" data-location="{location_filter}">'
        card_link_end = '</div>'
        actions_html = '<button class="btn-closed" disabled>✓ ' + ('Let Out' if 'let out' in ads_lower else 'Sold') + '</button>'
    else:
        card_link_start = f'<a href="{prop_id}.html" class="property-card" data-type="{type_filter}" data-location="{location_filter}">'
        card_link_end = '</a>'
        actions_html = f'<span class="btn-view">View Details →</span>'
    
    return f'''
            {card_link_start}
                <div class="property-image">
                    {image_html}
                    <div class="property-badges">
                        {status_badge}
                        <span class="badge badge-type {type_class}">{type_badge}</span>
                    </div>
                </div>
                <div class="property-content">
                    <p class="property-location">{location}</p>
                    <h3 class="property-type">{property_type}</h3>
                    <div class="property-specs">
                        <span>{beds_display}</span>
                        <span>🚿 {baths_str} Baths</span>
                        <span>📐 {sqft_str} sqft</span>
                    </div>
                    <p class="property-price">{price_html}</p>
                    <div class="property-actions">
                        {actions_html}
                    </div>
                </div>
            {card_link_end}'''


# =============================================================================
# GENERATE DETAIL PAGE
# =============================================================================

def generate_detail_page(row, output_folder):
    prop_id = row.get('Property_ID', '')
    location = row.get('Location', 'Unknown Location')
    property_type = row.get('Property type', 'Property')
    listing_type = row.get('Listing_Type', 'Rent')
    beds = row.get('Bedrooms', 0)
    baths = row.get('Bathrooms', 0)
    sqft = row.get('Built-up', 0)
    rent_price = row.get('Rent_RM', 0) if not pd.isna(row.get('Rent_RM', 0)) else row.get('Rental price', 0)
    sale_price = row.get('Sale_RM', 0) if not pd.isna(row.get('Sale_RM', 0)) else row.get('Selling price', 0)
    furnishing = row.get('Furnishing', 'N/A')
    tenure = row.get('Tenure', 'N/A')
    tour_link = row.get('3D_link', '')
    video_link = row.get('Video_Link', '')  # Ready for when you add this column
    property_folder = row.get('property_folder', '')
    
    # Read SEO description from seo.docx
    description = read_seo_docx(property_folder)
    if not description:
        description = f"Beautiful {property_type.lower()} located in {location}. Contact Adelyn for more details and to arrange a viewing."
    
    photos = copy_property_photos(property_folder, prop_id, output_folder, MAX_PHOTOS_DETAIL_PAGE)
    
    # Format values
    if pd.isna(beds):
        beds_str = '0'
    elif str(beds).lower() == 'studio':
        beds_str = 'Studio'
    else:
        try:
            beds_str = str(int(beds))
        except:
            beds_str = str(beds)
    
    baths_str = str(int(baths)) if not pd.isna(baths) and str(baths).replace('.','').isdigit() else '0'
    sqft_str = f"{int(sqft):,}" if not pd.isna(sqft) and str(sqft).replace('.','').replace(',','').isdigit() and float(sqft) > 0 else "N/A"
    
    if listing_type == 'Sale' and not pd.isna(sale_price) and sale_price > 0:
        price = f"RM {int(sale_price):,}"
        price_label = ""
        badge_class = "badge-sale"
        listing_type_display = "For Sale"
    else:
        price = f"RM {int(rent_price):,}" if not pd.isna(rent_price) and rent_price > 0 else "Contact"
        price_label = "per month"
        badge_class = "badge-rent"
        listing_type_display = "For Rent"
    
    # WhatsApp link
    wa_message = f"Hi Adelyn, I'm interested in the {property_type} at {location}"
    whatsapp_link = f"https://wa.me/{WHATSAPP}?text={wa_message.replace(' ', '%20')}"
    
    # Tour button (from 3D_link column)
    tour_button = ""
    if tour_link and not pd.isna(tour_link) and str(tour_link).strip():
        tour_button = f'<a href="{tour_link}" class="action-btn btn-tour" target="_blank">🔘 View 360° Tour</a>'
    
    # Video button (ready for future)
    video_button = ""
    if video_link and not pd.isna(video_link) and str(video_link).strip():
        video_button = f'<a href="{video_link}" class="action-btn btn-video" target="_blank">▶️ Watch Video</a>'
    
    # Tour embed section
    tour_section = ""
    if tour_link and not pd.isna(tour_link) and str(tour_link).strip():
        tour_section = f'''
        <div class="tour-section">
            <h2>360° Virtual Tour</h2>
            <div class="tour-embed">
                <iframe src="{tour_link}" allowfullscreen></iframe>
            </div>
        </div>'''
    
    # Hero images
    if len(photos) >= 3:
        hero_image = f"photos/{prop_id}/{photos[0]}"
        side_image_1 = f"photos/{prop_id}/{photos[1]}"
        side_image_2 = f"photos/{prop_id}/{photos[2]}"
        more_photos = len(photos) - 3
    elif len(photos) == 2:
        hero_image = f"photos/{prop_id}/{photos[0]}"
        side_image_1 = f"photos/{prop_id}/{photos[1]}"
        side_image_2 = f"photos/{prop_id}/{photos[0]}"
        more_photos = 0
    elif len(photos) == 1:
        hero_image = f"photos/{prop_id}/{photos[0]}"
        side_image_1 = f"photos/{prop_id}/{photos[0]}"
        side_image_2 = f"photos/{prop_id}/{photos[0]}"
        more_photos = 0
    else:
        hero_image = ""
        side_image_1 = ""
        side_image_2 = ""
        more_photos = 0
    
    # Gallery images
    gallery_images = ""
    for i, photo in enumerate(photos):
        gallery_images += f'<img src="photos/{prop_id}/{photo}" alt="{location}" onclick="openLightbox({i})" loading="lazy">'
    
    photos_json = ", ".join([f'"photos/{prop_id}/{p}"' for p in photos])
    
    meta_description = f"{property_type} for {listing_type_display.lower()} in {location}. {beds_str} bedrooms, {baths_str} bathrooms, {sqft_str} sqft."
    
    # Build HTML
    html = get_detail_template()
    html = html.replace('{property_title}', f"{property_type} at {location}")
    html = html.replace('{meta_description}', meta_description)
    html = html.replace('{hero_image}', hero_image)
    html = html.replace('{side_image_1}', side_image_1)
    html = html.replace('{side_image_2}', side_image_2)
    html = html.replace('{more_photos}', str(more_photos))
    html = html.replace('{property_type}', property_type)
    html = html.replace('{location}', location)
    html = html.replace('{beds}', beds_str)
    html = html.replace('{baths}', baths_str)
    html = html.replace('{sqft}', sqft_str)
    html = html.replace('{price}', price)
    html = html.replace('{price_label}', price_label)
    html = html.replace('{badge_class}', badge_class)
    html = html.replace('{listing_type}', listing_type_display)
    html = html.replace('{whatsapp_link}', whatsapp_link)
    html = html.replace('{tour_button}', tour_button)
    html = html.replace('{video_button}', video_button)
    html = html.replace('{description}', description)
    html = html.replace('{furnishing}', str(furnishing) if not pd.isna(furnishing) else 'N/A')
    html = html.replace('{tenure}', str(tenure) if not pd.isna(tenure) else 'N/A')
    html = html.replace('{tour_section}', tour_section)
    html = html.replace('{gallery_images}', gallery_images)
    html = html.replace('{photos_json}', photos_json)
    
    output_file = os.path.join(output_folder, f"{prop_id}.html")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    
    return output_file


# =============================================================================
# MAIN
# =============================================================================

def generate_website():
    print("=" * 60)
    print("ADELYN WONG WEBSITE GENERATOR v4")
    print("WITH PROPERTY DETAIL PAGES + SEO FROM seo.docx")
    print("=" * 60)
    
    if not os.path.exists(MASTERBOOK_PATH):
        print(f"\n❌ ERROR: Cannot find masterbook at:")
        print(f"   {MASTERBOOK_PATH}")
        input("\nPress Enter to exit...")
        return
    
    if not HAS_DOCX:
        print("\n⚠️ Installing python-docx...")
        import subprocess
        subprocess.run(['pip', 'install', 'python-docx', '--break-system-packages'], capture_output=True)
        print("   Please run the script again.")
        input("\nPress Enter to exit...")
        return
    
    if not HAS_PIL:
        print("\n⚠️ Installing Pillow for photo compression...")
        import subprocess
        subprocess.run(['pip', 'install', 'Pillow', '--break-system-packages'], capture_output=True)
        print("   Please run the script again.")
        input("\nPress Enter to exit...")
        return
    
    if COMPRESS_PHOTOS:
        print(f"\n📷 Photo compression: ENABLED")
        print(f"   Max size: {MAX_PHOTO_WIDTH}x{MAX_PHOTO_HEIGHT}, Quality: {JPEG_QUALITY}%")
    
    output_folder = os.path.abspath(OUTPUT_FOLDER)
    os.makedirs(os.path.join(output_folder, 'photos'), exist_ok=True)
    
    print(f"\n📖 Reading masterbook...")
    
    try:
        props = pd.read_excel(MASTERBOOK_PATH, sheet_name='Properties')
        active = pd.read_excel(MASTERBOOK_PATH, sheet_name='Active Listing')
        print(f"   ✅ Properties: {len(props)} rows")
        print(f"   ✅ Active Listing: {len(active)} rows")
        
        try:
            commercial = pd.read_excel(MASTERBOOK_PATH, sheet_name='Commercial_Properties')
            print(f"   ✅ Commercial: {len(commercial)} rows")
        except:
            commercial = pd.DataFrame()
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        input("\nPress Enter to exit...")
        return
    
    print("\n🔗 Merging data...")
    residential_active = active[active['Category'] == 'Residential']
    commercial_active = active[active['Category'] == 'Commercial']
    
    merged_residential = residential_active.merge(props, on='Property_ID', how='left', suffixes=('_listing', '_prop'))
    
    if len(commercial_active) > 0 and len(commercial) > 0:
        commercial_active_copy = commercial_active.copy()
        commercial_active_copy['Commercial_ID'] = commercial_active_copy['Property_ID']
        merged_commercial = commercial_active_copy.merge(commercial, on='Commercial_ID', how='left', suffixes=('_listing', '_prop'))
        merged = pd.concat([merged_residential, merged_commercial], ignore_index=True)
    else:
        merged = merged_residential
    
    print(f"   ✅ Total: {len(merged)} listings")
    
    available = merged[~merged['Ads_Status'].str.lower().str.contains('let out|sold', na=False)]
    closed = merged[merged['Ads_Status'].str.lower().str.contains('let out|sold', na=False)]
    
    total_properties = len(props) + len(commercial)
    active_count = len(available)
    closed_count = len(closed)
    
    print(f"\n📊 Stats: {active_count} available, {closed_count} closed")
    
    # Generate cards
    print("\n📷 Generating property cards...")
    property_cards = []
    
    for _, row in available.iterrows():
        card = generate_property_card(row, output_folder)
        property_cards.append(card)
    
    for _, row in closed.head(10).iterrows():
        card = generate_property_card(row, output_folder)
        property_cards.append(card)
    
    print(f"   ✅ {len(property_cards)} cards")
    
    # Generate detail pages
    print("\n📄 Generating detail pages (with SEO from seo.docx)...")
    detail_count = 0
    seo_count = 0
    for _, row in available.iterrows():
        property_folder = row.get('property_folder', '')
        seo_path = os.path.join(PHOTO_BASE_PATH, str(property_folder), 'seo.docx') if property_folder else ''
        if seo_path and os.path.exists(seo_path):
            seo_count += 1
        generate_detail_page(row, output_folder)
        detail_count += 1
    print(f"   ✅ {detail_count} detail pages")
    print(f"   ✅ {seo_count} with seo.docx found")
    
    # Build main page
    print("\n📝 Building main page...")
    html = get_main_template()
    html = html.replace('{whatsapp}', WHATSAPP)
    html = html.replace('{wechat_id}', WECHAT_ID)
    html = html.replace('{total_properties}', str(total_properties))
    html = html.replace('{active_listings}', str(active_count))
    html = html.replace('{closed_deals}', str(closed_count))
    html = html.replace('{property_cards}', '\n'.join(property_cards))
    html = html.replace('{generated_date}', datetime.now().strftime('%d %b %Y'))
    
    with open(os.path.join(output_folder, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ SUCCESS!")
    print(f"   📄 index.html")
    print(f"   📄 {detail_count} property pages (PROP-XXXXX.html)")
    print(f"   📷 photos/ folder (compressed to ~{JPEG_QUALITY}% quality)")
    print("\n" + "=" * 60)
    print("Next: Run Upload_To_GitHub.bat")
    print("=" * 60)
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    generate_website()
