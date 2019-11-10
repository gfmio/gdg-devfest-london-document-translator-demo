import clsx from "clsx";
import React, { FC, useState } from "react";
import IconButton from "@material-ui/core/IconButton";
import SnackbarContent from "@material-ui/core/SnackbarContent";
import { makeStyles } from "@material-ui/core/styles";
import CloseIcon from "@material-ui/icons/Close";
import ErrorIcon from "@material-ui/icons/Error";

interface ErrorMessageProps {
  message: string;
  closeable?: boolean;
  closed?: boolean;
  onClose?: () => void;
}

const useStyles = makeStyles((theme) => ({
  errorMessage: {
    margin: theme.spacing(1),
    backgroundColor: theme.palette.error.dark,
  },
  icon: {
    fontSize: 20,
  },
  iconVariant: {
    opacity: 0.9,
    marginRight: theme.spacing(1),
  },
  message: {
    display: "flex",
    alignItems: "center",
  },
}));

const ErrorMessage: FC<ErrorMessageProps> = ({ closeable, closed, onClose, message }) => {
  const classes = useStyles({});

  if (closed) {
    return null;
  }

  return (
    <SnackbarContent
      className={classes.errorMessage}
      message={
        <span id="client-snackbar" className={classes.message}>
          <ErrorIcon className={clsx(classes.icon, classes.iconVariant)} />
          {message}
        </span>
      }
      action={
        closeable
          ? [
              <IconButton key="close" aria-label="close" color="inherit" onClick={onClose}>
                <CloseIcon className={classes.icon} />
              </IconButton>,
            ]
          : undefined
      }
    />
  );
};

ErrorMessage.defaultProps = {
  closeable: false,
  closed: false,
  onClose: () => {},
};

export default ErrorMessage;
