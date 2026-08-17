"""
Generate a professional, multi-resolution icon for Sanjivani Attendance System.
Creates both icon.png and icon.ico (256, 128, 64, 48, 32, 16 px).
"""

from PIL import Image, ImageDraw

def create_app_icon():
    size = 512
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 1. Background rounded container with modern subtle gradient
    # Top-left navy (#0F172A) to bottom-right royal blue (#1E3A8A)
    corner_radius = 90
    draw.rounded_rectangle(
        [(24, 24), (size - 24, size - 24)],
        radius=corner_radius,
        fill="#0F172A",
        outline="#38BDF8",
        width=10
    )

    # Inner subtle glow rectangle
    draw.rounded_rectangle(
        [(40, 40), (size - 40, size - 40)],
        radius=corner_radius - 12,
        fill=None,
        outline="#1E293B",
        width=4
    )

    # 2. AI Face Scanner HUD Brackets (Cyan #38BDF8)
    bracket_color = "#38BDF8"
    b_len = 70
    b_thick = 14
    b_margin = 100

    # Top-Left Bracket
    draw.line([(b_margin, b_margin), (b_margin + b_len, b_margin)], fill=bracket_color, width=b_thick)
    draw.line([(b_margin, b_margin), (b_margin, b_margin + b_len)], fill=bracket_color, width=b_thick)

    # Top-Right Bracket
    draw.line([(size - b_margin, b_margin), (size - b_margin - b_len, b_margin)], fill=bracket_color, width=b_thick)
    draw.line([(size - b_margin, b_margin), (size - b_margin, b_margin + b_len)], fill=bracket_color, width=b_thick)

    # Bottom-Left Bracket
    draw.line([(b_margin, size - b_margin), (b_margin + b_len, size - b_margin)], fill=bracket_color, width=b_thick)
    draw.line([(b_margin, size - b_margin), (b_margin, size - b_margin - b_len)], fill=bracket_color, width=b_thick)

    # Bottom-Right Bracket
    draw.line([(size - b_margin, size - b_margin), (size - b_margin - b_len, size - b_margin)], fill=bracket_color, width=b_thick)
    draw.line([(size - b_margin, size - b_margin), (size - b_margin, size - b_margin - b_len)], fill=bracket_color, width=b_thick)

    # 3. Biometric Face Silhouette in center
    cx, cy = size // 2, size // 2 - 10
    
    # Head / Face Oval
    face_w, face_h = 85, 110
    draw.ellipse(
        [(cx - face_w, cy - face_h), (cx + face_w, cy + face_h)],
        fill="#2563EB",
        outline="#60A5FA",
        width=6
    )

    # Shoulders / Bust
    draw.chord(
        [(cx - 140, cy + 60), (cx + 140, cy + 240)],
        start=180, end=360,
        fill="#1D4ED8",
        outline="#3B82F6",
        width=6
    )

    # AI Recognition Scan Line across face
    scan_color = "#34D399"  # Emerald green scan line
    draw.line([(cx - 110, cy), (cx + 110, cy)], fill=scan_color, width=8)

    # Facial Landmark Points (AI Nodes)
    points = [
        (cx - 38, cy - 25), (cx + 38, cy - 25), # Eyes
        (cx, cy + 15),                          # Nose
        (cx - 30, cy + 55), (cx + 30, cy + 55), # Mouth corners
        (cx, cy + 65)                           # Chin
    ]
    for px, py in points:
        draw.ellipse([(px - 7, py - 7), (px + 7, py + 7)], fill="#F8FAFC", outline="#38BDF8", width=3)

    # 4. Verified Checkmark Badge at bottom-right
    badge_cx, badge_cy = size - 115, size - 115
    badge_r = 52
    draw.ellipse(
        [(badge_cx - badge_r, badge_cy - badge_r), (badge_cx + badge_r, badge_cy + badge_r)],
        fill="#16A34A",
        outline="#DCFCE7",
        width=8
    )

    # White Checkmark
    check_pts = [
        (badge_cx - 24, badge_cy + 2),
        (badge_cx - 6, badge_cy + 20),
        (badge_cx + 26, badge_cy - 16)
    ]
    draw.line(check_pts[:2], fill="#FFFFFF", width=10)
    draw.line(check_pts[1:], fill="#FFFFFF", width=10)

    # Save high-res PNGs
    img.save("sanjivani.png", "PNG")
    img.save("icon.png", "PNG")
    print("Saved sanjivani.png and icon.png")

    # Generate multi-size ICO (256, 128, 64, 48, 32, 16 px)
    icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    
    # Save sanjivani.ico and icon.ico with all standard resolutions
    img.save("sanjivani.ico", format="ICO", sizes=icon_sizes)
    img.save("icon.ico", format="ICO", sizes=icon_sizes)
    print(f"Saved sanjivani.ico and icon.ico with sizes: {icon_sizes}")

if __name__ == "__main__":
    create_app_icon()

