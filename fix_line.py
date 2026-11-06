with open('kapituluak/03_errepidea.tex', 'r', encoding='utf-8') as f:
    lines = f.readlines()

correct = 'Simulazioan, $0{,}30$~m/s-ko gehieneko abiadurarekin, bost biren batez bestekoa 40 segundo ingurukoa izan zen. RKA6rekin gehieneko abiadura $0{,}60$~m/s-ra igo genuen, eta bost biren batez bestekoa $20{,}82$~s izan zen. Pista errealean bira azkarrago osatzeko, zuzenean azkarrago joateaz gain, bihurgunetik nola ateratzen zen doitu behar izan genuen.\n'
lines[117] = correct
with open('kapituluak/03_errepidea.tex', 'w', encoding='utf-8') as f:
    f.writelines(lines)
print('OK:', repr(lines[117][:60]))
