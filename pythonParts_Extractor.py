import csv
from collections import OrderedDict

# داده‌های ورودی
inspection_data = [
    # SPH-08
    ['SPH-08', '2', 'HIPPS has no support.', 'Pi', '', 'P'],
    ['SPH-08', '2', 'The piping is installed in the AG / UG section obliquely.', 'Pi', '', 'P'],
    ['SPH-08', '2', 'All DBBs are unoperated.', 'Pi', '', 'P'],
    ['SPH-08', '2', 'The supports beyond the concrete pads have settled down.', 'Pi', '', 'P'],
    ['SPH-08', '2', 'valve 4" after SST & All DBB Valves need the painting touch up', 'Co', '', 'P'],
    
    # SPH-09
    ['SPH-09', '2', 'U-bolts and rubbers are not perfect for CIP pipelines.', 'Pi', '', 'P'],
    ['SPH-09', '2', 'The face flange is in direct contact with the temporary spool, it may cause damage to the face flange.', 'Pi', '', 'P'],
    ['SPH-09', '2', 'The paint of HIPPS both of valves is peeled off', 'Co', '', 'P'],
    ['SPH-09', '2', 'TG002 not installed.', 'Pi', '', 'P'],
    ['SPH-09', '2', 'Sec A has not pressure gauge.', 'Pi', '', 'P'],
    ['SPH-09', '2', 'SS valve used in S.C line (piping Sec C)', 'Pi', '', 'P'],
    ['SPH-09', '2', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-09', '2', 'piping of daily diesel drum is unpainted.', 'Co', '', 'P'],
    ['SPH-09', '2', 'The 2-inch line of the annulus does not have a final layer of paint.', 'Co', '', 'P'],
    ['SPH-09', '2', 'The joint weld of the annulus line is not painted.', 'Co', '', 'P'],
    ['SPH-09', '2', 'Support piping 2" Sec C not fixed and need to add support under valve 1".', 'Pi', '', 'P'],
    
    # JR-06
    ['JR-06', '2', 'TG0002 is not installed.', 'Pi', '', 'P'],
    ['JR-06', '2', 'The burn pit line has shifted from the support at the end before exiting the fence and needs a guide installed.', 'Pi', '', 'P'],
    ['JR-06', '2', 'pressure gauge on the Sec B has not been installed.', 'Pi', '', 'P'],
    ['JR-06', '2', 'The 6" valve before the STT has peeling paint in some areas.', 'Pi', '', 'P'],
    ['JR-06', '2', 'PSLL001 is not installed.', 'Pi', '', 'P'],
    ['JR-06', '2', 'All DBBVs & IJ need touch-up.', 'Pi', '', 'P'],
    ['JR-06', '2', 'All valves at the end of the piping are rusty and need touch-up.', 'Pi', '', 'P'],
    ['JR-06', '2', 'Flange 4" after STT has leakage.', 'Pi', '', 'P'],
    ['JR-06', '2', 'Flange 6" before HIPPS has leakage.', 'Pi', '', 'P'],
    ['JR-06', '2', 'The piping is positioned 3 centimeters from the first support.', 'Pi', '', 'P'],
    ['JR-06', '2', 'The coating on the joint after the IJ needs repair.', 'Co', '', 'P'],
    
    # JR-05
    ['JR-05', '2', 'The burn pit line is dented and misaligned in several places', 'Pi', '', 'P'],
    ['JR-05', '2', 'The slope of the burn pit line has not been maintained.', 'Pi', '', 'P'],
    ['JR-05', '2', 'The platform frame of access is not compliant with the standards and is not approved.', 'Pi', '', 'P'],
    ['JR-05', '2', 'The gratings are welded together, which is not compliant with the standards.', 'Pi', '', 'P'],
    ['JR-05', '2', 'None of the supports have been grouted.', 'Pi', '', 'P'],
    
    # NM
    ['NM', '2', 'All globe valves are jammed and unusable.', 'Pi', '', 'P'],
    ['NM', '2', 'All DBB valves are jammed and unusable.', 'Pi', '', 'P'],
    ['NM', '2', 'The test header line has not been painted.', 'Co', '', 'P'],
    ['NM', '2', 'The test header line has not been connected and is not operational.', 'Pi', '', 'P'],
    ['NM', '2', 'All ESDVs are not equipped with tubing & panels.', 'Pi', '', 'P'],
    ['NM', '2', 'The bypass valve line 2" CRD-120-00008-B7C3NRX-PP has not an actuator.', 'Pi', '', 'P'],
    ['NM', '2', 'The scaffolding under the Jufair 7 inlet pipeline is still present from the construction time.', 'Pi', '', 'P'],
    ['NM', '2', 'The 8-inch IK is painted, which causes it to be ineffective.', 'Co', '', 'P'],
    
    # JR-08
    ['JR-08', '2', 'The chemical package has not been completed and is not started.', 'Pi', '', 'P'],
    ['JR-08', '2', 'All Manual valves must be lubricated.', 'Pi', '', 'P'],
    ['JR-08', '2', 'All manual valves are locked as per P&ID and do not have LOTO tag.', 'Pi', '', 'P'],
    ['JR-08', '2', 'All Valves have experienced paint splashes and damage to their finish due to inadequate covering during the painting process.', 'Co', '', 'P'],
    ['JR-08', '2', 'has not been painted properly overall and has various defects.', 'Co', '', 'P'],
    
    # SM-02
    ['SM-02', '2', 'The oil tanks of the ESDV have no support', 'Pi', '', 'P'],
    ['SM-02', '2', 'The supports beyond the concrete pads have settled down.', 'Pi', '', 'P'],
    ['SM-02', '2', 'All DBBs are unoperated.', 'Pi', '', 'P'],
    ['SM-02', '2', 'IK 2"&8" is not installed.', 'Pi', '', 'P'],
    ['SM-02', '2', 'All Close drain piping are unpainted.', 'Co', '', 'P'],
    ['SM-02', '2', 'piping of daily diesel drum is unpainted.', 'Co', '', 'P'],
    ['SM-02', '2', 'The inlet pipeline needs touching up coated.', 'Co', '', 'P'],
    
    # RA
    ['RA', '2', 'All the oil tanks of the ESDV have completely come off their supports (0003-0006, 0002-0005, 0004-0007)', 'Pi', '', 'P'],
    ['RA', '2', 'U-Bolts lack rubber.', 'Pi', '', 'P'],
    ['RA', '2', 'The inlet piping from the north manifold is not fully installed and the support is not complete.', 'Pi', '', 'P'],
    ['RA', '2', 'The 2" drain line is not connected to the 18" drain line.', 'Pi', '', 'P'],
    ['RA', '2', 'IK related to input lines SM 1, SM 2 and Nm some washers are broken.', 'Pi', '', 'P'],
    ['RA', '2', 'ALL DBB are unoperated.', 'Pi', '', 'P'],
    ['RA', '2', 'Close drain piping is unpainted.', 'Co', '', 'P'],
    ['RA', '2', 'piping of daily diesel drum is unpainted.', 'Co', '', 'P'],
    ['RA', '2', 'Equalizer line 2" of ESDV0007 are unpainted.', 'Co', '', 'P'],
    
    # SPH-06
    ['SPH-06', '', 'U-bolts and rubbers are not perfect for CIP pipelines.', 'Pi', '', 'P'],
    ['SPH-06', '', 'The flanges of the CIP lines are not assembled.', 'Pi', '', 'P'],
    ['SPH-06', '', 'HIPPS has no support.', 'Pi', '', 'P'],
    ['SPH-06', '', 'The Anchor block of the pipeline has sunk.', 'Pi', '', 'P'],
    ['SPH-06', '', 'The spools before and after barred tee is distortion (Due The support type PS08 for line 6-CRD-100-56037 has not installed).', 'Pi', '', 'P'],
    ['SPH-06', '', 'CDH Piping is not final painted.', 'Co', '', 'P'],
    ['SPH-06', '', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-06', '', 'ALL piping of daily diesel drum is unpainted.', 'Co', '', 'P'],
    ['SPH-06', '', 'The piping is installed in the AG / UG section obliquely.', 'Pi', '', 'P'],
    ['SPH-06', '', 'All DBBs are unoperated.', 'Pi', '', 'P'],
    ['SPH-06', '', 'The supports beyond the concrete pads have settled down.', 'Pi', '', 'P'],
    ['SPH-06', '', 'The face flange is in direct contact with the temporary spool, it may cause damage to the face flange.', 'Pi', '', 'P'],
    ['SPH-06', '', 'TG002 not installed.', 'Pi', '', 'P'],
    ['SPH-06', '', 'Sec A has not pressure gauge.', 'Pi', '', 'P'],
    ['SPH-06', '', 'SS valve used in S.C line (piping Sec C)', 'Pi', '', 'P'],
    ['SPH-06', '', 'The paint of HIPPS both of valves is peeled off', 'Co', '', 'P'],
    ['SPH-06', '', 'The 2-inch line of the annulus does not have a final layer of paint.', 'Co', '', 'P'],
    ['SPH-06', '', 'The joint weld of the annulus line are not painted.', 'Co', '', 'P'],
    
    # SPH-11
    ['SPH-11', '', 'All DBB Valves are not operated.', 'Pi', '', 'P'],
    ['SPH-11', '', 'PG TOP has not installed.', 'Pi', '', 'P'],
    ['SPH-11', '', 'CIP Package & Piping has not been installed.', 'Pi', '', 'P'],
    ['SPH-11', '', 'The second support after X mass is located at a distance of 5 centimeters.', 'Pi', '', 'P'],
    ['SPH-11', '', 'Hot Bend pipe & IJ need to be touched up', 'Co', '', 'P'],
    
    # SPH-02
    ['SPH-02', '', 'HIPPS haven`t support.', 'Pi', '', 'P'],
    ['SPH-02', '', 'The support at the end of the piping has settled down.', 'Pi', '', 'P'],
    ['SPH-02', '', 'Line 2" CRD-100-06052-B10C3NR-PP has not completed and do not have Support.', 'Pi', '', 'P'],
    ['SPH-02', '', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-02', '', 'piping of daily diesel drum is unpainted.', 'Co', '', 'P'],
    ['SPH-02', '', 'Flange before HIPPS is not painted.', 'Co', '', 'P'],
    
    # SPH-03
    ['SPH-03', '', 'There is no 4-inch blind, instead, there is a l injection nozzle and stainless stee valve with a 2-inch blind.', 'Pi', '', 'P'],
    ['SPH-03', '', 'The stainless-steel valve installed on the line is dangerous due to the possibility of galvanic corrosion of the valve body and related parts.', 'Pi', '', 'P'],
    ['SPH-03', '', 'PSLL 0003 is not installed.', 'Pi', '', 'P'],
    ['SPH-03', '', 'HIPPS haven`t support.', 'Pi', '', 'P'],
    ['SPH-03', '', 'CIP piping is not re-installed after HIPPS installation operation and is left in non-standard form.', 'Pi', '', 'P'],
    ['SPH-03', '', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-03', '', 'piping of daily diesel drum is unpainted.', 'Co', '', 'P'],
    ['SPH-03', '', 'There is galvanic corrosion in the of the nut& bolt & stem of CIP Valve.', 'Co', '', 'P'],
    ['SPH-03', '', 'Flange before HIPPS is not painted.', 'Co', '', 'P'],
    ['SPH-03', '', 'The Paint of the 90-degree elbow after STT has been worn.', 'Co', '', 'P'],
    
    # SPH-04
    ['SPH-04', '2', 'PSLL 0003 is not installed', 'Pi', '', 'P'],
    ['SPH-04', '2', 'The CIP lines do not have proper support, and the rubber under the SS piping is not available.', 'Pi', '', 'P'],
    ['SPH-04', '2', 'The piping for the daily diesel drum is incomplete.', 'Pi', '', 'P'],
    ['SPH-04', '2', 'The support piping 4" after STT has settled down.', 'Pi', '', 'P'],
    ['SPH-04', '2', 'DBB of PG 003 & 001A,C HIPPS`s PT is not operable.', 'Pi', '', 'P'],
    ['SPH-04', '2', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-04', '2', 'piping of daily diesel drum are unpainted.', 'Co', '', ''],
    ['SPH-04', '2', 'There is galvanic corrosion in the of the nut& bolt & stem of CIP Valve.', 'Co', '', ''],
    
    # SPH-05
    ['SPH-05', '', 'TG 002 is not Installe.', 'Pi', '', 'P'],
    ['SPH-05', '', 'The CIP lines do not have proper support, and the rubber under the SS piping is not available.', 'Pi', '', 'P'],
    ['SPH-05', '', 'PSHH02, PSLL03 are not installe.', 'Pi', '', 'P'],
    ['SPH-05', '', 'The 2" CRD valves and piping lack support and guide brackets.', 'Pi', '', 'P'],
    ['SPH-05', '', 'The HIPPS does not have suitable support.', 'Pi', '', 'P'],
    ['SPH-05', '', 'The support at the end of the piping  has settled down.', 'Co', '', 'P'],
    ['SPH-05', '', 'piping of daily diesel drum are unpainted.', 'Co', '', 'P'],
    ['SPH-05', '', 'There is galvanic corrosion in the of the nut& bolt & stem of CIP Valve.', 'Co', '', 'P'],
    ['SPH-05', '', 'Some parts of the HIPPS shelter are not painted.', 'Co', '', 'P'],
    ['SPH-05', '', 'The flanges of the CIP lines are not assembled.', 'Pi', '', 'P'],
    ['SPH-05', '', 'CDH Piping is not final painted.', 'Pi', '', 'P'],
    ['SPH-05', '', 'Hipps has no support.', 'Pi', '', 'P'],
    ['SPH-05', '', 'The Anchor block of the pipeline has sunk.', 'Pi', '', 'P'],
    ['SPH-05', '', 'The spools before and after bardtee is distortion (Due The support type PS08 for line 6-CRD-100-56037 has not installed).', 'Co', '', 'P'],
    ['SPH-05', '', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-05', '', 'ALL piping of daily diesel drum are unpainted.', 'Co', '', 'P'],
    ['SPH-05', '', 'All CDH Piping is unpainted final layer.', 'Co', '', 'P'],
    
    # SPH-07
    ['SPH-07', '', 'Hipps has no support', 'Pi', '', 'P'],
    ['SPH-07', '', 'Valve`s handwheel before STT has difficulty in operation.', 'Pi', '', 'P'],
    ['SPH-07', '', 'The paint of HIPPS valve is peeled off', 'Pi', '', 'P'],
    ['SPH-07', '', 'Sec A has not peressure gauge.', 'Pi', '', 'P'],
    ['SPH-07', '', 'SS valve used in S.C line(piping Sec C)', 'Pi', '', 'P'],
    ['SPH-07', '', 'The paint of HIPPS both of valves is peeled off', 'Co', '', 'P'],
    ['SPH-07', '', 'All CIP pipes are unpainted.', 'Co', '', 'P'],
    ['SPH-07', '', 'The piping is installed in the AG / UG section obliquely.', 'Pi', '', 'P'],
    ['SPH-07', '', 'All DBBs are unoperated.', 'Pi', '', 'P'],
    ['SPH-07', '', 'The supports beyond the concrete pads has settled down.', 'Pi', '', 'P'],
]

# حذف موارد تکراری
seen = set()
unique_data = []
duplicates_removed = 0

for row in inspection_data:
    # ایجاد کلید یکتا بر اساس Location و Inspection Issues
    key = (row[0], row[2])
    if key not in seen:
        seen.add(key)
        unique_data.append(row)
    else:
        duplicates_removed += 1

# مرتب‌سازی بر اساس Location
unique_data.sort(key=lambda x: x[0])

# ذخیره در فایل CSV
output_file = 'Piping_Inspection_Cleaned.csv'
with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    # نوشتن هدر
    writer.writerow(['Location', 'Inspection', 'Inspection Issues', 'Discipline', 'Production Shortage', 'Production Punch'])
    # نوشتن داده‌ها
    writer.writerows(unique_data)

print(f"✅ فایل CSV با موفقیت ایجاد شد: {output_file}")
print(f"📊 تعداد کل موارد قبل از حذف تکراری: {len(inspection_data)}")
print(f"📊 تعداد موارد یکتا: {len(unique_data)}")
print(f"🗑️ تعداد موارد حذف شده: {duplicates_removed}")

# نمایش 10 ردیف اول
print("\n📋 نمونه داده‌ها (10 ردیف اول):")
print("-" * 150)
print(f"{'Location':<15} {'Insp':<6} {'Issue':<80} {'Disc':<6} {'Shortage':<10} {'Punch':<6}")
print("-" * 150)
for i, row in enumerate(unique_data[:10], 1):
    issue = row[2][:77] + "..." if len(row[2]) > 80 else row[2]
    print(f"{row[0]:<15} {row[1]:<6} {issue:<80} {row[3]:<6} {row[4]:<10} {row[5]:<6}")
print("-" * 150)
