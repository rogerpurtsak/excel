"""Tests for SQLite job-state persistence."""

from src.job_manager import JobManager, JobState


def _mgr(tmp_path):
    return JobManager(db_path=tmp_path / "jobs.db")


def test_create_and_get_job(tmp_path):
    mgr = _mgr(tmp_path)
    jid = mgr.create_job("/path/book.xls", "immediate")
    job = mgr.get_job(jid)
    assert job["workbook_path"] == "/path/book.xls"
    assert job["state"] == JobState.PENDING.value
    mgr.close()


def test_active_job_lookup(tmp_path):
    mgr = _mgr(tmp_path)
    jid = mgr.create_job("/p/b.xls", "gradual")
    assert mgr.get_active_job("/p/b.xls")["id"] == jid
    mgr.complete_job(jid)
    assert mgr.get_active_job("/p/b.xls") is None
    mgr.close()


def test_progress_and_state_transitions(tmp_path):
    mgr = _mgr(tmp_path)
    jid = mgr.create_job("/p/b.xls", "gradual")
    mgr.update_job_progress(jid, 2, 5)
    job = mgr.get_job(jid)
    assert job["completed_tasks"] == 2 and job["total_tasks"] == 5
    assert job["state"] == JobState.RUNNING.value

    mgr.pause_job(jid)
    assert mgr.get_job(jid)["state"] == JobState.PAUSED.value
    mgr.resume_job(jid)
    assert mgr.get_job(jid)["state"] == JobState.RUNNING.value
    mgr.cancel_job(jid)
    assert mgr.get_job(jid)["state"] == JobState.CANCELLED.value
    mgr.close()


def test_save_and_load_state_blob(tmp_path):
    mgr = _mgr(tmp_path)
    jid = mgr.create_job("/p/b.xls", "immediate")
    mgr.save_job_state(jid, {"completed": ["t1", "t2"], "cursor": 2})
    loaded = mgr.load_job_state(jid)
    assert loaded == {"completed": ["t1", "t2"], "cursor": 2}
    mgr.close()
