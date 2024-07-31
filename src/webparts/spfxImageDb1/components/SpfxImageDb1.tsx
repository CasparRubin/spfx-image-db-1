import * as React from "react";
import styles from "./SpfxImageDb1.module.scss";
import type { ISpfxImageDb1Props } from "./ISpfxImageDb1Props";
import { makeStyles, Spinner } from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    "> div": { padding: "20px" },
  },
});

const SpfxImageDb1: React.FC<ISpfxImageDb1Props> = (props) => {
  const { hasTeamsContext } = props;

  const fluentStyles = useStyles();

  return (
    <section
      className={`${styles.spfxImageDb1} ${
        hasTeamsContext ? styles.teams : ""
      }`}
    >
      <div className={fluentStyles.container}>
        <Spinner labelPosition="above" label="Label Position Above..." />
      </div>
    </section>
  );
};

export default SpfxImageDb1;