interface ConfirmDialogProps {
  open: boolean;
  title: string;
  message: string;
  danger?: boolean;
  confirmLabel?: string;
  cancelLabel?: string;
  onConfirm: () => void;
  onCancel: () => void;
}

export default function ConfirmDialog({ open, title, message, danger, confirmLabel = 'Confirm', cancelLabel = 'Cancel', onConfirm, onCancel }: ConfirmDialogProps) {
  return (
    <div className={`acc-modal-overlay${open ? ' on' : ''}`} onClick={onCancel}>
      <div className="acc-modal" role="alertdialog" aria-modal="true" aria-labelledby="acc-modal-title" onClick={(e) => e.stopPropagation()}>
        <h2 id="acc-modal-title">{title}</h2>
        <p>{message}</p>
        <div className="acc-modal-actions">
          <button type="button" className="acc-btn" onClick={onCancel}>
            {cancelLabel}
          </button>
          <button type="button" className={`acc-btn ${danger ? 'acc-btn-danger' : 'acc-btn-primary'}`} onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
