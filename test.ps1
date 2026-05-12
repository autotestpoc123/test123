.no-texting-popover {
  --bs-popover-max-width: 592px;
  border: none;
  border-radius: 6px;
  box-shadow: 0 6px 24px rgba(0, 0, 0, 0.15);
  background: #fff;
  max-width: calc(100vw - 32px);
}

.no-texting-popover .popover-arrow,
.no-texting-popover::before,
.no-texting-popover::after {
  display: none !important;
}

.no-texting-popover .popover-body {
  padding: 0;
  color: inherit;
}
