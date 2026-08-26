"""P6 ilerleme girisi kurallari -- veritabanisiz birim testleri.

Burada test edilen sey "UPDATE calisti mi" degil, P6 semantiginin korunup
korunmadigi. Bir status guncellemesinde tutarsiz kalan alan, programi bozuk
degil SESSIZCE YANLIS yapar:

* status_code fiili tarihlerle uyusmazsa P6 aktiviteyi bambaska gosterir;
* kalan sure sifirlanmadan aktivite "bitti" olursa sonraki reschedule biten
  isin arkasina is iter;
* kaynak yuklu bir aktivitede atamanin kalan birimi guncellenmezse P6
  reschedule sirasinda kalan sureyi geri yazar (bukhtourcity85'te olculdu:
  72 saat yazildi, F9 sonrasi 240 saate dondu).
"""
import datetime as _dt

import pytest

from p6 import progress


def task(**kw):
    base = {
        "task_id": 1, "task_code": "A001", "task_name": "Test",
        "status_code": progress.STATUS_NOT_STARTED,
        "complete_pct_type": "CP_Drtn",
        "phys_complete_pct": 0, "target_drtn_hr_cnt": 80,
        "remain_drtn_hr_cnt": 80, "act_start_date": None, "act_end_date": None,
        "target_work_qty": 100, "act_work_qty": 0, "remain_work_qty": 100,
    }
    base.update(kw)
    return base


DD = _dt.datetime(2026, 11, 1)


def plan(update, cur=None, data_date=DD, allow_future=False):
    return progress._plan(update, cur or task(), data_date, allow_future)


# --- tamamlanan aktivite ---------------------------------------------------
def test_complete_tum_alanlari_tutarli_birakir():
    f = plan({"status": "complete", "actual_start": "2026-09-02",
              "actual_finish": "2026-09-24"})
    assert f["status_code"] == progress.STATUS_COMPLETE
    assert f["remain_drtn_hr_cnt"] == 0.0
    assert f["phys_complete_pct"] == 100.0
    assert f["remain_work_qty"] == 0.0
    assert f["act_work_qty"] == 100
    assert f["act_start_date"] == _dt.datetime(2026, 9, 2)
    assert f["act_end_date"] == _dt.datetime(2026, 9, 24)


def test_complete_fiili_bitis_yoksa_reddedilir():
    with pytest.raises(progress.ProgressError, match="actual_finish"):
        plan({"status": "complete", "actual_start": "2026-09-02"})


def test_complete_kayitli_baslangici_kullanir():
    cur = task(act_start_date=_dt.datetime(2026, 9, 2),
               status_code=progress.STATUS_ACTIVE)
    f = plan({"status": "complete", "actual_finish": "2026-09-24"}, cur)
    assert f["act_start_date"] == _dt.datetime(2026, 9, 2)


def test_bitis_baslangictan_once_olamaz():
    with pytest.raises(progress.ProgressError, match="once olamaz"):
        plan({"status": "complete", "actual_start": "2026-09-24",
              "actual_finish": "2026-09-02"})


# --- devam eden aktivite ---------------------------------------------------
def test_yuzde_kalan_sureye_cevrilir():
    f = plan({"status": "in_progress", "actual_start": "2026-09-18",
              "percent_complete": 70})
    assert f["status_code"] == progress.STATUS_ACTIVE
    assert f["remain_drtn_hr_cnt"] == pytest.approx(24.0)   # 80h * 0.30
    assert f["phys_complete_pct"] == 70.0
    assert f["act_end_date"] is None


def test_kalan_sure_yuzdeye_cevrilir():
    f = plan({"status": "in_progress", "actual_start": "2026-09-18",
              "remaining_duration_h": 20})
    assert f["remain_drtn_hr_cnt"] == 20.0
    assert f["phys_complete_pct"] == pytest.approx(75.0)    # (80-20)/80


def test_devam_edende_yuzde_veya_kalan_sure_sart():
    with pytest.raises(progress.ProgressError, match="percent_complete veya"):
        plan({"status": "in_progress", "actual_start": "2026-09-18"})


def test_devam_eden_yuzde_100_olamaz():
    with pytest.raises(progress.ProgressError, match="status='complete'"):
        plan({"status": "in_progress", "actual_start": "2026-09-18",
              "percent_complete": 100})


def test_devam_edene_fiili_bitis_verilemez():
    with pytest.raises(progress.ProgressError, match="complete"):
        plan({"status": "in_progress", "actual_start": "2026-09-18",
              "percent_complete": 50, "actual_finish": "2026-10-01"})


def test_baslangic_yoksa_baslatilamaz():
    with pytest.raises(progress.ProgressError, match="actual_start"):
        plan({"status": "in_progress", "percent_complete": 50})


def test_negatif_kalan_sure_reddedilir():
    with pytest.raises(progress.ProgressError, match="negatif"):
        plan({"status": "in_progress", "actual_start": "2026-09-18",
              "remaining_duration_h": -5})


# --- baslamamis ------------------------------------------------------------
def test_not_started_her_seyi_geri_alir():
    cur = task(status_code=progress.STATUS_COMPLETE, phys_complete_pct=100,
               remain_drtn_hr_cnt=0, act_start_date=_dt.datetime(2026, 9, 2),
               act_end_date=_dt.datetime(2026, 9, 24), act_work_qty=100,
               remain_work_qty=0)
    f = plan({"status": "not_started"}, cur)
    assert f["act_start_date"] is None and f["act_end_date"] is None
    assert f["remain_drtn_hr_cnt"] == 80
    assert f["remain_work_qty"] == 100
    assert f["phys_complete_pct"] == 0.0


# --- durum cikarimi --------------------------------------------------------
def test_status_verilmezse_tarihten_cikarilir():
    assert plan({"actual_finish": "2026-09-24", "actual_start": "2026-09-02"}
                )["status_code"] == progress.STATUS_COMPLETE
    assert plan({"actual_start": "2026-09-02", "percent_complete": 40}
                )["status_code"] == progress.STATUS_ACTIVE


def test_turkce_ve_ingilizce_status_takma_adlari():
    for alias in ("complete", "bitti", "TK_Complete", "COMPLETED"):
        f = plan({"status": alias, "actual_start": "2026-09-02",
                  "actual_finish": "2026-09-24"})
        assert f["status_code"] == progress.STATUS_COMPLETE
    for alias in ("in_progress", "devam", "active"):
        f = plan({"status": alias, "actual_start": "2026-09-02",
                  "percent_complete": 10})
        assert f["status_code"] == progress.STATUS_ACTIVE


def test_gecersiz_status_reddedilir():
    with pytest.raises(progress.ProgressError, match="gecersiz"):
        plan({"status": "yarim", "actual_start": "2026-09-02"})


# --- veri tarihi -----------------------------------------------------------
def test_veri_tarihinden_sonraki_fiil_reddedilir():
    cur = task()
    f = plan({"status": "complete", "actual_start": "2026-11-05",
              "actual_finish": "2026-11-20"})
    with pytest.raises(progress.ProgressError, match="veri tarihinden"):
        progress._check_dates(f, cur, DD, allow_future=False)


def test_allow_future_ile_uyariya_dusulur():
    cur = task()
    f = plan({"status": "complete", "actual_start": "2026-11-05",
              "actual_finish": "2026-11-20"})
    warn = progress._check_dates(f, cur, DD, allow_future=True)
    assert warn and "veri tarihinden" in warn[0]


def test_veri_tarihi_yoksa_uyari_verilir():
    warn = progress._check_dates(plan({"status": "complete",
                                       "actual_start": "2026-09-02",
                                       "actual_finish": "2026-09-24"}),
                                 task(), None, False)
    assert warn and "denetlenemedi" in warn[0]


# --- tarih ayristirma ------------------------------------------------------
@pytest.mark.parametrize("value,expected", [
    ("2026-09-24", _dt.datetime(2026, 9, 24)),
    ("2026-09-24 17:00", _dt.datetime(2026, 9, 24, 17, 0)),
    ("2026-09-24T17:00:00", _dt.datetime(2026, 9, 24, 17, 0)),
    ("24.09.2026", _dt.datetime(2026, 9, 24)),
    (None, None),
])
def test_tarih_formatlari(value, expected):
    assert progress._parse_date(value, "t") == expected


def test_bozuk_tarih_reddedilir():
    with pytest.raises(progress.ProgressError, match="okunamadi"):
        progress._parse_date("dun", "actual_start")


@pytest.mark.parametrize("bad", [-1, 101, "cok"])
def test_gecersiz_yuzde_reddedilir(bad):
    with pytest.raises(progress.ProgressError):
        progress._pct(bad, "percent_complete")


# --- atama senkronu --------------------------------------------------------
def _asg(target_qty=240, rate=5.0, target_cost=1200):
    return {"taskrsrc_id": 7, "task_id": 1, "target_qty": target_qty,
            "remain_qty": target_qty, "act_reg_qty": 0,
            "target_cost": target_cost, "act_reg_cost": 0,
            "remain_cost": target_cost, "cost_per_qty": rate}


def test_atama_yuzdeye_gore_bolunur():
    (asg_id, act_qty, remain_qty, act_cost, remain_cost), = \
        progress._assignment_plan([_asg()], 70.0)
    assert asg_id == 7
    assert act_qty == pytest.approx(168.0)     # 240 * 0.70
    assert remain_qty == pytest.approx(72.0)
    assert act_cost == pytest.approx(840.0)    # 168 * 5
    assert remain_cost == pytest.approx(360.0)


def test_tamamlanan_atamanin_kalani_sifirlanir():
    _id, act_qty, remain_qty, _ac, remain_cost = \
        progress._assignment_plan([_asg()], 100.0)[0]
    assert remain_qty == 0.0 and remain_cost == 0.0
    assert act_qty == 240.0


def test_ucret_yoksa_maliyet_hedeften_bolunur():
    _id, _aq, _rq, act_cost, remain_cost = \
        progress._assignment_plan([_asg(rate=0, target_cost=1000)], 25.0)[0]
    assert act_cost == pytest.approx(250.0)
    assert remain_cost == pytest.approx(750.0)


def test_atamasiz_aktivite_bos_plan_dondurur():
    assert progress._assignment_plan([], 50.0) == []


# --- fark raporu -----------------------------------------------------------
def test_diff_yalnizca_degisenleri_gosterir():
    cur = task()
    f = plan({"status": "in_progress", "actual_start": "2026-09-18",
              "percent_complete": 50})
    delta = progress._diff(cur, f)
    assert "status_code" in delta and "act_start_date" in delta
    assert delta["status_code"] == {"from": "TK_NotStart", "to": "TK_Active"}


def test_degisiklik_yoksa_diff_bos():
    cur = task()
    assert progress._diff(cur, plan({"status": "not_started"}, cur)) == {}
