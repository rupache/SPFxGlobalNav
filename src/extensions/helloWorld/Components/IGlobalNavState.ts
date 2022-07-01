import { IOrderedTermInfo } from '@pnp/sp/taxonomy';
export interface IGobalNavState {
    loading: boolean;
    terms: IOrderedTermInfo[];
}