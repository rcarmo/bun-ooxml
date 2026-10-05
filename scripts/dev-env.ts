// Make/package entrypoint preload, never loaded by runtime APIs.
import {prepareDevPaths} from './dev-paths.ts';
prepareDevPaths();
