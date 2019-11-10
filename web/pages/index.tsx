import download from "in-browser-download";
import React, { FC, useState, useEffect } from "react";
import Button from "@material-ui/core/Button";
import Card from "@material-ui/core/Card";
import CardContent from "@material-ui/core/CardContent";
import CircularProgress from "@material-ui/core/CircularProgress";
import FormControl from "@material-ui/core/FormControl";
import Grid from "@material-ui/core/Grid";
import InputLabel from "@material-ui/core/InputLabel";
import MenuItem from "@material-ui/core/MenuItem";
import Select from "@material-ui/core/Select";
import Typography from "@material-ui/core/Typography";
import { makeStyles } from "@material-ui/core/styles";
import Api, { Language } from "../api";
import ErrorMessage from "../components/ErrorMessage";
import Layout from "../components/Layout";
import newFileName from "../util/newFileName";

/** Represents a "virtual file" consisting of a name and buffer */
interface VirtualFile {
  buffer: ArrayBuffer;
  name: string;
}

interface State {
  error: any;
  errorMessageClosed: boolean;
  supportedLanguages: Language[] | null;
  from: string | null;
  to: string;
  isTranslating: boolean;
  file: File | null;
  translatedFile: VirtualFile | null;
}

const defaultState = {
  error: null,
  errorMessageClosed: true,
  from: null,
  isTranslating: false,
  to: "es",
  file: null,
  translatedFile: null,
};

const useStyles = makeStyles((theme) => ({
  card: {
    maxWidth: "33em",
    margin: theme.spacing(3, 2),
  },
  formControl: {
    margin: theme.spacing(1),
    minWidth: 120,
    width: "100%",
  },
  input: {
    display: "none",
  },
  button: {
    margin: theme.spacing(1),
  },
}));

/** Word, Excel and PowerPoint documents */
const acceptedMimeTypes = [
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
];

const App: FC<{}> = () => {
  // Styles
  const classes = useStyles({});

  // State
  const [state, setEntireState] = useState<State>({
    ...defaultState,
    supportedLanguages: null,
  });
  const setState = (partialState: Partial<State>) => setEntireState({ ...state, ...partialState });
  const resetState = () => setState(defaultState);
  const { error, errorMessageClosed, file, from, isTranslating, supportedLanguages, to, translatedFile } = state;

  const canSubmit = !!supportedLanguages && !!file && !isTranslating;

  // Request the supported languages on load
  useEffect(() => {
    (async () => {
      try {
        const supportedLanguages = await Api.getSupportedLanguages();
        setState({ supportedLanguages });
      } catch (e) {
        setState({ error: e });
      }
    })();
  }, []);

  // Event handlers
  const onFromChange = (e: React.ChangeEvent<{ name?: string; value: string | null }>) => {
    e.preventDefault();
    setState({ from: e.target.value || null });
  };

  const onToChange = (e: React.ChangeEvent<{ name?: string; value: string }>) => {
    e.preventDefault();
    setState({ to: e.target.value });
  };

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    setState({ file: e.target.files[0] || null });
  };

  const onSubmitClick = async (e: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
    e.preventDefault();

    if (!canSubmit) {
      return;
    }

    // Set isTranslating to true, remove the translated file and any previous error
    setState({
      isTranslating: true,
      translatedFile: null,
      error: null,
      errorMessageClosed: true,
    });

    try {
      const name = newFileName(file, to, from);

      const buffer = await Api.translateDocument(file, to, from);

      setState({
        isTranslating: false,
        translatedFile: {
          buffer,
          name,
        },
      });
    } catch (e) {
      setState({
        isTranslating: false,
        error: e,
        errorMessageClosed: false,
      });
    }
  };

  const onResetClick = (e: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
    e.preventDefault();
    resetState();
  };

  const onDownloadTranslatedDocumentClick = (e: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
    e.preventDefault();
    if (!translatedFile) {
      return;
    }
    download(translatedFile.buffer, translatedFile.name);
  };

  const onErrorMessageClose = () => setState({ errorMessageClosed: true });

  // Render

  // Display an error if one occurred while fetching the supported languages or when there are none
  if ((error && !supportedLanguages) || (supportedLanguages && supportedLanguages.length === 0)) {
    return (
      <Layout>
        <Card className={classes.card}>
          <CardContent>
            <ErrorMessage message="An error occurred while requesting the supported languages." />;
          </CardContent>
        </Card>
      </Layout>
    );
  }

  // Display a spinner while fetching the supported languages
  if (!supportedLanguages) {
    return (
      <Layout>
        <Card className={classes.card}>
          <CardContent style={{ textAlign: "center" }}>
            <Typography>Loading supported languages...</Typography>
            <CircularProgress />
          </CardContent>
        </Card>
      </Layout>
    );
  }

  // If all required data is available, display the form
  return (
    <Layout>
      <Card className={classes.card}>
        <CardContent>
          <form>
            <Grid container spacing={2}>
              <Grid item xs={6}>
                {/* From */}
                <FormControl className={classes.formControl}>
                  <InputLabel shrink id="from-label">
                    From
                  </InputLabel>
                  <Select displayEmpty={true} labelId="from-label" value={from || ""} onChange={onFromChange}>
                    <MenuItem key={-1} value="">
                      Detect automatically
                    </MenuItem>
                    {supportedLanguages.map(({ code, name }, index) => (
                      <MenuItem key={index} value={code}>
                        {name}
                      </MenuItem>
                    ))}
                  </Select>
                </FormControl>
              </Grid>
              <Grid item xs={6}>
                {/* To */}
                <FormControl className={classes.formControl}>
                  <InputLabel shrink id="to-label">
                    To
                  </InputLabel>
                  <Select labelId="to-label" value={to} onChange={onToChange}>
                    {supportedLanguages.map(({ code, name }, index) => (
                      <MenuItem key={index} value={code}>
                        {name}
                      </MenuItem>
                    ))}
                  </Select>
                </FormControl>
              </Grid>
              <Grid item xs={12}>
                {/* Document */}
                <input
                  accept={acceptedMimeTypes.join(",")}
                  className={classes.input}
                  id="file-input"
                  type="file"
                  onChange={onFileChange}
                />
                <label htmlFor="file-input">
                  <Button variant="contained" component="span" className={classes.button}>
                    Select file
                  </Button>
                  <Typography component="span" variant="caption">
                    {file ? `${file.name} (${file.size} bytes)` : "No file selected"}
                  </Typography>
                </label>
              </Grid>
              {/* Buttons */}
              <Grid item xs={12}>
                <Button
                  variant="contained"
                  color="primary"
                  className={classes.button}
                  onClick={onSubmitClick}
                  disabled={!canSubmit}
                >
                  {isTranslating ? <CircularProgress size="24px" /> : "Translate Document"}
                </Button>
                <Button className={classes.button} onClick={onResetClick}>
                  Reset
                </Button>
              </Grid>

              {/* Error message */}
              {error && (
                <Grid item xs={12}>
                  <ErrorMessage
                    closeable
                    closed={errorMessageClosed}
                    onClose={onErrorMessageClose}
                    message="An error occurred while translating the document."
                  />
                </Grid>
              )}
            </Grid>
          </form>

          <Grid container spacing={2}>
            <Grid item xs={12}>
              {/* Download */}
              <Button
                variant="contained"
                color="secondary"
                className={classes.button}
                disabled={!translatedFile}
                onClick={onDownloadTranslatedDocumentClick}
              >
                Download
              </Button>
              <Typography component="span" variant="caption">
                {translatedFile ? `${translatedFile.name} (${translatedFile.buffer.byteLength} bytes)` : ""}
              </Typography>
            </Grid>
          </Grid>
        </CardContent>
      </Card>
    </Layout>
  );
};

export default App;
