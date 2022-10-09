import { LoginComponent } from './login/login.component';
import { ProfileComponent } from './profile/profile.component';
import { PageNotFoundComponent } from './page-not-found/page-not-found.component';
import { DocumentListComponent } from './document-list/document-list.component';
import { HomeComponent } from './home/home.component';
import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { UserauthguardGuard } from './userauthguard.guard';

const routes: Routes = [
  { path: 'home', component: HomeComponent },
  {
    path: 'document',
    component: DocumentListComponent,
    canActivate: [UserauthguardGuard],
  },
  {
    path: 'profile',
    component: ProfileComponent,
    canActivate: [UserauthguardGuard],
  },
  { path: 'login', component: LoginComponent },
  { path: '', redirectTo: '/home', pathMatch: 'full' },
  { path: '**', component: PageNotFoundComponent },
];

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule],
})
export class AppRoutingModule {}
